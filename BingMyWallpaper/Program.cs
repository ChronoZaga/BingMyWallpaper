using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;

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
                string json;
                using (var client = new WebClient())
                {
                    json = client.DownloadString(bingApiUrl);
                }

                // Parse JSON using DataContractJsonSerializer
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    var serializer = new DataContractJsonSerializer(typeof(BingResponse));
                    var data = (BingResponse)serializer.ReadObject(ms);

                    if (data.Images == null || data.Images.Length == 0)
                    {
                        Console.WriteLine("No image data found in Bing API response.");
                        Environment.Exit(1); // Exit with error code
                    }

                    // Extract image URL and filename
                    string imageUrl = "https://www.bing.com" + data.Images[0].Url;
                    string filename = $"{data.Images[0].StartDate}-{data.Images[0].Title.Replace(" ", "-").Replace("?", "")}.jpg";
                    string wallpaperPath = Path.Combine(tempPath, filename);

                    // Download the 4K image
                    using (var client = new WebClient())
                    {
                        client.DownloadFile(imageUrl, wallpaperPath);
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
                        Console.WriteLine("Failed to download wallpaper.");
                        Environment.Exit(1); // Exit with error code
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Environment.Exit(1); // Exit with error code
            }
        }

        private static void SetWallpaper(string path)
        {
            // Set the wallpaper using Windows API
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        }
    }
}