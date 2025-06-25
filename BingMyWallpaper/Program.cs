using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;

namespace BingWallpaperDownloader
{
    class Program
    {
        // Windows API constants and declarations
        private const int SPI_SETDESKWALLPAPER = 20;
        private const int SPIF_UPDATEINIFILE = 0x01;
        private const int SPIF_SENDWININICHANGE = 0x02;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

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

        static void Main(string[] args)
        {
            try
            {
                // Determine days back (idx) from command-line argument
                int daysBack = 0;
                if (args.Length > 0 && int.TryParse(args[0], out int parsedDays))
                {
                    daysBack = Math.Max(0, Math.Min(parsedDays, 7)); // Limit to 0-7 days per Bing API
                }

                Console.WriteLine($"Changing wallpaper to Bing's Picture from {daysBack} day(s) ago...");

                // Get temp directory
                string tempPath = Path.GetTempPath();

                // Bing API URL for 4K resolution, with dynamic idx
                string bingApiUrl = $"https://www.bing.com/HPImageArchive.aspx?format=js&idx={daysBack}&n=1&mkt=en-US&uhd=1&uhdwidth=3840&uhdheight=2160";

                // Download JSON data
                string json = ""; // Initialize to avoid unassigned variable error
                try
                {
                    using (var client = new WebClient())
                    {
                        json = client.DownloadString(bingApiUrl);
                    }
                }
                catch (WebException webEx)
                {
                    string errorMessage = $"Failed to download Bing API data from {bingApiUrl}. Error: {webEx.Message}{(webEx.InnerException != null ? $" Inner Error: {webEx.InnerException.Message}" : "")}";
                    Console.WriteLine(errorMessage);
                    MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Environment.Exit(1);
                }

                // Parse JSON using DataContractJsonSerializer
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    var serializer = new DataContractJsonSerializer(typeof(BingResponse));
                    var data = (BingResponse)serializer.ReadObject(ms);

                    if (data.Images == null || data.Images.Length == 0)
                    {
                        string errorMessage = "No image data found in Bing API response.";
                        Console.WriteLine(errorMessage);
                        MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(1); // Exit with error code
                    }

                    // Extract image URL and filename
                    string imageUrl = "https://www.bing.com" + data.Images[0].Url;
                    string title = CleanFileName(data.Images[0].Title);
                    string filename = $"{data.Images[0].StartDate}-{title}.jpg";
                    string wallpaperPath = Path.Combine(tempPath, filename);

                    // Download the 4K image
                    try
                    {
                        using (var client = new WebClient())
                        {
                            client.DownloadFile(imageUrl, wallpaperPath);
                        }
                    }
                    catch (WebException webEx)
                    {
                        string errorMessage = $"Failed to download image from {imageUrl}. Error: {webEx.Message}{(webEx.InnerException != null ? $" Inner Error: {webEx.InnerException.Message}" : "")}";
                        Console.WriteLine(errorMessage);
                        MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(1);
                    }

                    // Verify the file exists
                    if (File.Exists(wallpaperPath))
                    {
                        // Set the wallpaper directly as JPG
                        SetWallpaper(wallpaperPath);
                        Console.WriteLine("Wallpaper set successfully!");
                        Environment.Exit(0); // Exit with success code
                    }
                    else
                    {
                        string errorMessage = "Failed to download wallpaper. File not found at " + wallpaperPath;
                        Console.WriteLine(errorMessage);
                        MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(1); // Exit with error code
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMessage = $"An error occurred: {ex.Message}{(ex.InnerException != null ? $" Inner Error: {ex.InnerException.Message}" : "")}";
                Console.WriteLine(errorMessage);
                MessageBox.Show(errorMessage, "Bing Wallpaper Downloader - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1); // Exit with error code
            }
        }

        private static string CleanFileName(string fileName)
        {
            // Replace invalid filename characters with a hyphen
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }
            // Replace spaces with hyphens for consistency
            fileName = fileName.Replace(" ", "-");
            // Ensure no trailing or leading hyphens
            return fileName.Trim('-');
        }

        private static void SetWallpaper(string path)
        {
            // Set the wallpaper using Windows API
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        }
    }
}