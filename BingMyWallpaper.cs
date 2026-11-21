// BingMyWallpaper - working copy
// Written for C# 5 so it builds with the compiler included in Windows
// (C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe). Build with "BingMyWallpaper Build.bat".
//
// Usage:
//   BingMyWallpaper.exe        Set today's Bing image as the wallpaper (%USERPROFILE%\wallpaper.jpg, hidden, Fill).
//   BingMyWallpaper.exe 0-7    Same, using the image from that many days ago (0 = today).
//   BingMyWallpaper.exe DL     Download the last 7 days of images to the desktop (wallpaper unchanged).
//
// Runs as a standard user; administrator rights are never needed.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace BingWallpaperDownloader
{
    class Program
    {
        // Where the wallpaper is stored (C:\Users\<you>\wallpaper.jpg). Overwritten on every run.
        private static readonly string WallpaperFile =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "wallpaper.jpg");

        // Images are downloaded here first, so a failed or partial download never replaces a good file.
        // Fail() deletes it, because Environment.Exit skips finally blocks.
        private static readonly string TempFile =
            Path.Combine(Path.GetTempPath(), "BingMyWallpaper-download.jpg");

        // At login the network is often not ready yet; retry Bing requests before reporting an error.
        private const int StartupAttempts = 6;
        private const int RetryDelayMs = 10000;

        // Windows API constants and declarations (used only if the IDesktopWallpaper API is unavailable)
        private const int SPI_SETDESKWALLPAPER = 20;
        private const int SPIF_UPDATEINIFILE = 0x01;
        private const int SPIF_SENDWININICHANGE = 0x02;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

        // IDesktopWallpaper (shobjidl.h) - the API the Settings app itself uses. Method order must match the header.
        [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDesktopWallpaper
        {
            void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);
            [return: MarshalAs(UnmanagedType.LPWStr)] string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID);
            [return: MarshalAs(UnmanagedType.LPWStr)] string GetMonitorDevicePathAt(uint monitorIndex);
            uint GetMonitorDevicePathCount();
            Rect GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID);
            void SetBackgroundColor(uint color);
            uint GetBackgroundColor();
            void SetPosition(DesktopWallpaperPosition position);
            DesktopWallpaperPosition GetPosition();
            void SetSlideshow(IntPtr items);
            IntPtr GetSlideshow();
            void SetSlideshowOptions(int options, uint slideshowTick);
            void GetSlideshowOptions(out int options, out uint slideshowTick);
            void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, int direction);
            int GetStatus();
            void Enable([MarshalAs(UnmanagedType.Bool)] bool enable);
        }

        [ComImport, Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD")]
        private class DesktopWallpaperClass { }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect { public int Left, Top, Right, Bottom; }

        private enum DesktopWallpaperPosition { Center = 0, Tile = 1, Stretch = 2, Fit = 3, Fill = 4, Span = 5 }

        // Data contract for JSON parsing
        [System.Runtime.Serialization.DataContract]
        private class BingResponse
        {
            [System.Runtime.Serialization.DataMember(Name = "images")]
            public ImageData[] Images { get; set; }
        }

        [System.Runtime.Serialization.DataContract]
        private class ImageData
        {
            [System.Runtime.Serialization.DataMember(Name = "url")]
            public string Url { get; set; }

            [System.Runtime.Serialization.DataMember(Name = "startdate")]
            public string StartDate { get; set; }

            [System.Runtime.Serialization.DataMember(Name = "title")]
            public string Title { get; set; }
        }

        [STAThread] // required for the IDesktopWallpaper COM object and MessageBox
        static void Main(string[] args)
        {
            // Exes built with the in-box compiler have no target-framework attribute and can
            // default to old TLS versions that Bing rejects, so require TLS 1.2 explicitly.
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            try
            {
                // "DL" argument downloads the last 7 days of wallpapers to the desktop
                if (args.Length > 0 && args[0].Equals("DL", StringComparison.OrdinalIgnoreCase))
                {
                    DownloadLastSevenDays();
                    Environment.Exit(0); // Exit with success code
                }

                // Determine days back (idx) from command-line argument
                int daysBack = 0;
                int parsedDays;
                if (args.Length > 0 && int.TryParse(args[0], out parsedDays))
                {
                    daysBack = Math.Max(0, Math.Min(parsedDays, 7)); // Limit to 0-7 days per Bing API
                }

                Console.WriteLine("Changing wallpaper to Bing's Picture from " + daysBack + " day(s) ago...");

                // Bing API URL for 4K resolution, with dynamic idx
                string bingApiUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=" + daysBack +
                                    "&n=1&mkt=en-US&uhd=1&uhdwidth=3840&uhdheight=2160";

                BingResponse data = GetBingData(bingApiUrl, StartupAttempts);
                string imageUrl = ImageUrl(data.Images[0]);
                if (imageUrl == null)
                {
                    Fail("Bing returned no image URL.");
                }

                // Download to the temp file first, so a failed or partial download never replaces
                // the current wallpaper with a broken file (which would leave a black desktop).
                string error;
                if (!TryDownloadImage(imageUrl, TempFile, StartupAttempts, out error))
                {
                    Fail("Failed to download wallpaper from " + imageUrl + ". Error: " + error);
                }

                ReplaceWallpaperFile(TempFile, WallpaperFile);
                TryDelete(TempFile);

                SetWallpaper(WallpaperFile);
                Console.WriteLine("Wallpaper set successfully!");
                Environment.Exit(0); // Exit with success code
            }
            catch (Exception ex)
            {
                Fail("An error occurred: " + ex.Message + InnerMessage(ex));
            }
        }

        // Saves the last 7 days of images to the desktop with their original names.
        // A failed image is skipped and reported at the end instead of stopping the rest.
        private static void DownloadLastSevenDays()
        {
            Console.WriteLine("Downloading Bing's Pictures from the last 7 days to your desktop...");

            // Get desktop directory
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            // Bing API URL for 4K resolution, today plus the 6 days prior
            string bingApiUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=7&mkt=en-US&uhd=1&uhdwidth=3840&uhdheight=2160";

            BingResponse data = GetBingData(bingApiUrl, 1);
            var failed = new List<string>();

            foreach (var image in data.Images)
            {
                // The date prefix (e.g. 20261121-) also keeps names clear of reserved names like CON or NUL.
                string filename = CleanFileName(image.StartDate) + "-" + CleanFileName(image.Title) + ".jpg";

                string imageUrl = ImageUrl(image);
                string error;
                if (imageUrl == null || !TryDownloadImage(imageUrl, TempFile, 1, out error))
                {
                    failed.Add(filename);
                    continue;
                }

                // Copy the finished download into place, so a partial file never lands on the desktop.
                try
                {
                    File.Copy(TempFile, Path.Combine(desktopPath, filename), true);
                }
                catch (Exception)
                {
                    failed.Add(filename);
                }
            }

            TryDelete(TempFile);

            if (failed.Count > 0)
            {
                Fail("Couldn't download " + failed.Count + " of " + data.Images.Length + " images:\n" + string.Join("\n", failed));
            }

            Console.WriteLine("Wallpapers downloaded successfully!");
        }

        // Returns the full image URL, or null if Bing didn't provide one.
        private static string ImageUrl(ImageData image)
        {
            if (image == null || string.IsNullOrEmpty(image.Url)) return null;
            return "https://www.bing.com" + image.Url;
        }

        // Copies the downloaded image over the wallpaper file and marks it hidden.
        private static void ReplaceWallpaperFile(string source, string destination)
        {
            // Windows refuses to overwrite a hidden or read-only file, so clear those attributes first.
            if (File.Exists(destination))
            {
                File.SetAttributes(destination, FileAttributes.Normal);
            }

            File.Copy(source, destination, true);
            File.SetAttributes(destination, FileAttributes.Hidden);
        }

        // Downloads and parses the Bing API JSON; exits with an error if it fails or has no images.
        private static BingResponse GetBingData(string bingApiUrl, int attempts)
        {
            // Raw bytes, not DownloadString: on .NET Framework that decodes with the ANSI code page,
            // which garbles accented titles. The JSON serializer reads the UTF-8 bytes correctly.
            byte[] json = null;
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    using (var client = new WebClient())
                    {
                        json = client.DownloadData(bingApiUrl);
                    }
                    break;
                }
                catch (WebException webEx)
                {
                    if (attempt < attempts)
                    {
                        Thread.Sleep(RetryDelayMs); // network may still be connecting after login
                        continue;
                    }
                    Fail("Failed to download Bing API data from " + bingApiUrl + ". Error: " + webEx.Message + InnerMessage(webEx));
                }
            }

            using (var ms = new MemoryStream(json))
            {
                var serializer = new DataContractJsonSerializer(typeof(BingResponse));
                var data = (BingResponse)serializer.ReadObject(ms);

                if (data == null || data.Images == null || data.Images.Length == 0)
                {
                    Fail("No image data found in Bing API response.");
                }
                return data;
            }
        }

        // Downloads an image, retrying up to the given number of attempts. Never leaves a partial file behind.
        private static bool TryDownloadImage(string imageUrl, string path, int attempts, out string error)
        {
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    using (var client = new WebClient())
                    {
                        client.DownloadFile(imageUrl, path);
                    }
                    if (new FileInfo(path).Length > 0)
                    {
                        error = null;
                        return true;
                    }
                    error = "The downloaded file was empty.";
                }
                catch (WebException webEx)
                {
                    error = webEx.Message + InnerMessage(webEx);
                }

                TryDelete(path);
                if (attempt >= attempts) return false;
                Thread.Sleep(RetryDelayMs); // network may have dropped briefly
            }
        }

        private static void TryDelete(string path)
        {
            try { File.Delete(path); } catch { }
        }

        private static string InnerMessage(Exception ex)
        {
            return ex.InnerException != null ? " Inner Error: " + ex.InnerException.Message : "";
        }

        private static void Fail(string errorMessage)
        {
            TryDelete(TempFile); // Environment.Exit skips finally blocks, so clean up here
            Console.WriteLine(errorMessage);
            MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1); // Exit with error code
        }

        // Turns a Bing title into a safe Windows file name that uses only characters on a US keyboard.
        // Example: Cote d'Azur with accents, curly quotes and an en dash -> "Cote-d'Azur-Sunset-Nice"
        private static string CleanFileName(string fileName)
        {
            // Split accented letters into letter + accent mark (e-acute -> e + accent), then drop the marks.
            string decomposed = (fileName ?? "").Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            foreach (char c in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(KeyboardText(c));
                }
            }

            // Spaces become hyphens, and runs of hyphens collapse to one.
            string name = Regex.Replace(sb.ToString().Replace(' ', '-'), "-{2,}", "-");

            // Keep names short so the full path stays well under the 260-character Windows limit.
            if (name.Length > MaxTitleLength) name = name.Substring(0, MaxTitleLength);

            // Windows drops a trailing period or space, so trim those along with stray hyphens.
            name = name.Trim('-', '.', ' ');
            return name.Length > 0 ? name : "Bing";
        }

        private const int MaxTitleLength = 100;

        // Maps one character to text that is valid in a Windows file name and typeable on a US keyboard.
        private static string KeyboardText(char c)
        {
            switch (c)
            {
                // Characters Windows forbids in file names
                case '<': case '>': case ':': case '"': case '/': case '\\': case '|': case '?': case '*':
                    return "-";
                // Letters that don't split into letter + accent
                // (escaped so the source stays plain ASCII for the in-box compiler)
                case '\u00DF': return "ss"; // sharp s
                case '\u00E6': return "ae";
                case '\u00C6': return "AE";
                case '\u0153': return "oe";
                case '\u0152': return "OE";
                case '\u00F8': return "o";  // slashed o
                case '\u00D8': return "O";
                case '\u0142': return "l";  // Polish l
                case '\u0141': return "L";
                case '\u0111': case '\u00F0': return "d";
                case '\u0110': case '\u00D0': return "D";
                case '\u00FE': return "th"; // thorn
                case '\u00DE': return "Th";
                case '\u0131': return "i";  // dotless i
                // Curly quotes, dashes and other typography
                case '\u2018': case '\u2019': case '\u201A': case '\u2032': return "'";
                case '\u201C': case '\u201D': case '\u201E': case '\u00AB': case '\u00BB': return "";
                case '\u2010': case '\u2011': case '\u2012': case '\u2013': case '\u2014': case '\u2015': return "-";
                case '\u2026': return "...";
                case '\u00A0': return " "; // non-breaking space
            }

            // Plain printable ASCII is on the keyboard; anything else becomes a hyphen.
            return (c >= ' ' && c <= '~') ? c.ToString() : "-";
        }

        // Sets the wallpaper on all monitors in Fill mode.
        private static void SetWallpaper(string path)
        {
            // Fill = WallpaperStyle 10, TileWallpaper 0. Written so the setting sticks for the fallback path too.
            using (RegistryKey desktop = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true))
            {
                if (desktop != null)
                {
                    desktop.SetValue("WallpaperStyle", "10");
                    desktop.SetValue("TileWallpaper", "0");
                }
            }

            // Use IDesktopWallpaper, the same API the Settings app uses. Unlike SystemParametersInfo,
            // it switches Windows out of Slideshow mode, so Windows' background state stays consistent
            // instead of conflicting with a slideshow (which made Settings > Background slow to load).
            IDesktopWallpaper wallpaper = null;
            try
            {
                wallpaper = (IDesktopWallpaper)new DesktopWallpaperClass();
                wallpaper.SetPosition(DesktopWallpaperPosition.Fill);
                wallpaper.SetWallpaper(null, path); // null = all monitors
            }
            catch (Exception)
            {
                SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
            }
            finally
            {
                if (wallpaper != null) Marshal.ReleaseComObject(wallpaper);
            }
        }
    }
}
