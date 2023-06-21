# Create the WMD folder if it doesn't exist
$destinationPath = "C:\ProgramData\WMD"
if (!(Test-Path -Path $destinationPath)) {
    New-Item -ItemType Directory -Path $destinationPath
}

# Copy all files and folders in the script's directory to the WMD folder
$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path
Get-ChildItem -Path $scriptDirectory -Recurse | ForEach-Object {
    $destinationFile = Join-Path -Path $destinationPath -ChildPath $_.FullName.Substring($scriptDirectory.length)
    if (!(Test-Path -Path (Split-Path -Path $destinationFile -Parent))) {
        New-Item -ItemType Directory -Path (Split-Path -Path $destinationFile -Parent)
    }
    Copy-Item -Path $_.FullName -Destination $destinationFile
}

# Create a new task in the Task Scheduler to run a bat file called RunLater.bat
$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c `"$destinationPath\RunLater.bat`"" -WorkingDirectory $destinationPath
$trigger = New-ScheduledTaskTrigger -Daily -At 2pm
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "MyTask" -Description "Runs my script daily at 9 AM"

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

# Check if Outlook is already running
$proc = Get-Process | Where-Object { $_.Name -eq "OUTLOOK" }

# Create an instance of Outlook
$outlook = New-Object -ComObject Outlook.Application

# Create a new Mail Item
$mail = $outlook.CreateItem(0)

# Set properties of the Mail Item
$mail.To = "pfuller@global-business.net"#"ktischler@global-business.net; lprevost@global-business.net; kwright@global-business.net; jfrancoeur@global-business.net; bmedcalf@global-business.net; ncasey@global-business.net; kmeade@global-business.net; pfuller@global-business.net; rrobertson@global-business.net ;de.robertson@global-business.net"#"bmedcalf@global-business.net"
$mail.Subject = "Keep Your Computer Locked"
$mail.Body = Get-Content ".\TheGreatestIntern.txt"

# Send the email
$mail.Send()

# Wait for a bit to allow the send process to complete
Start-Sleep -s 5

# Quit Outlook only if it was not running before the script started
if ($proc -eq $null) {
    $outlook.Quit()
}
