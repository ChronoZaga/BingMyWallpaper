# BingMyWallpaper

A small Windows app that sets your desktop wallpaper to the Bing Picture of the Day in 4K.

- No installer and no admin rights. It runs as a standard user.
- Nothing runs in the background. It sets the wallpaper and exits.
- Works with Windows 10 and 11, using the .NET Framework that's already part of Windows.

## Usage

| Command | What it does |
|---|---|
| `BingMyWallpaper.exe` | Sets today's Bing image as your wallpaper. |
| `BingMyWallpaper.exe 3` | Sets the image from 3 days ago. Use any number from `0` (today) to `7`. |
| `BingMyWallpaper.exe DL` | Downloads the last 7 days of images to your Desktop. Your wallpaper isn't changed. |

## How it works

**Setting the wallpaper**
- The image is saved as `wallpaper.jpg` in your user folder (`C:\Users\yourname\wallpaper.jpg`). The file is hidden and replaced on each run, so old images don't pile up.
- It's set on all monitors in **Fill** mode, through the same Windows API the Settings app uses. If your background was set to Slideshow, it's switched to Picture.
- The image is downloaded to a temp file first. If the download fails, your current wallpaper stays as it is.
- If the network isn't ready yet, for example right after you sign in, it retries for about a minute before showing an error.

**Downloading with `DL`**
- Files are named with the date and the image title, for example `20261121-Cote-d'Azur-Sunset-Nice.jpg`.
- Titles are cleaned up so names are valid on Windows and use only characters on a US keyboard. Accents are removed (é → e), curly quotes and long dashes are replaced, and other symbols become hyphens.
- If an image fails to download, the others still download, and a message at the end lists the ones that failed.

**Errors**
- Any error shows a short message box. The app returns exit code `0` on success and `1` on failure.

## Run it automatically at sign-in

1. Press **Win + R**, type `shell:startup` and press Enter.
2. Put a shortcut to `BingMyWallpaper.exe` in the folder that opens.

You get a fresh wallpaper every time you sign in.

## Building from source

The app is a single C# file built with the C# compiler included with Windows, so you don't need Visual Studio or an SDK.

1. Put `BingMyWallpaper.cs`, `BMW.ico` and `BingMyWallpaper Build.bat` in the same folder.
2. Run `BingMyWallpaper Build.bat`.

The script builds `BingMyWallpaper.exe` as a 64-bit (x64) app with the icon.

## License

MIT. See [LICENSE](LICENSE).
