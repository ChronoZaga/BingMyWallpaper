@echo off
setlocal
rem Builds BingMyWallpaper.exe from BingMyWallpaper.cs using the C# compiler included with Windows (.NET Framework 4.x).
cd /d "%~dp0"

set "VERSION=1.5.0.0"

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo ERROR: C# compiler not found at "%CSC%"
    pause
    exit /b 1
)

rem Generate the version attributes in a temp file so nothing extra is left on the desktop.
set "VERFILE=%TEMP%\BingMyWallpaper_Version.cs"
> "%VERFILE%" echo using System.Reflection;
>>"%VERFILE%" echo [assembly: AssemblyTitle("BingMyWallpaper")]
>>"%VERFILE%" echo [assembly: AssemblyProduct("BingMyWallpaper")]
>>"%VERFILE%" echo [assembly: AssemblyVersion("%VERSION%")]
>>"%VERFILE%" echo [assembly: AssemblyFileVersion("%VERSION%")]
>>"%VERFILE%" echo [assembly: AssemblyInformationalVersion("%VERSION%")]

echo Building BingMyWallpaper.exe version %VERSION%...
"%CSC%" /nologo /target:winexe /platform:x64 /optimize+ ^
    /win32icon:BMW.ico ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Runtime.Serialization.dll ^
    /out:BingMyWallpaper.exe ^
    BingMyWallpaper.cs "%VERFILE%"
set "RESULT=%ERRORLEVEL%"
del "%VERFILE%" 2>nul

if not "%RESULT%"=="0" (
    echo.
    echo BUILD FAILED
    pause
    exit /b 1
)

echo.
echo Build succeeded: %~dp0BingMyWallpaper.exe (version %VERSION%)
pause
