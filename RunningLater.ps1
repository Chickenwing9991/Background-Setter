# The rest of the script remains the same
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

# Get a list of all files in the folder
$files = Get-ChildItem -Path '.\Backgrounds' -Include *.jpg,*.jpeg,*.png,*.bmp -Recurse

# Select a random file
$randomFile = Get-Random -InputObject $files

[Wallpaper]::SystemParametersInfo(20, 0, $randomFile.FullName, 3)