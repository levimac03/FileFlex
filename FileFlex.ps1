# FileFlex - Professional File Conversion Tool
# Created on March 05, 2025

# Set professional console appearance
$Host.UI.RawUI.BackgroundColor = "DarkGray"
$Host.UI.RawUI.ForegroundColor = "White"
$Host.UI.RawUI.WindowTitle = "FileFlex - File Conversion Utility"
Clear-Host

# Function to display a professional header
function Show-Header {
    Clear-Host
    Write-Host "┌────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│           FileFlex Conversion Tool             │" -ForegroundColor White
    Write-Host "├────────────────────────────────────────────────┤" -ForegroundColor Gray
    Write-Host "│ Drag and drop a file or enter its path below.  │" -ForegroundColor White
    Write-Host "│ Type 'exit' to quit, 'help' for instructions.  │" -ForegroundColor White
    Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
    Write-Host ""
}

# Function to display help
function Show-Help {
    Clear-Host
    Write-Host "┌────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│                  FileFlex Help                 │" -ForegroundColor White
    Write-Host "├────────────────────────────────────────────────┤" -ForegroundColor Gray
    Write-Host "│ 1. Drag a file into this window or type its    │" -ForegroundColor White
    Write-Host "│    full path (e.g., C:\file.txt).              │" -ForegroundColor White
    Write-Host "│ 2. Select a conversion option by number.       │" -ForegroundColor White
    Write-Host "│ 3. Converted file will appear in the same      │" -ForegroundColor White
    Write-Host "│    directory as the original.                  │" -ForegroundColor White
    Write-Host "│ 4. Type 'back' to return, 'exit' to quit.      │" -ForegroundColor White
    Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Press Enter to return"
}

# Supported file types and conversion options (unchanged for brevity)
$conversionOptions = @{
    ".txt" = @("CSV", "JSON", "XML", "HTML")
    ".csv" = @("JSON", "XML", "XLSX", "TXT")
    # ... (rest of the options as before)
}

# Function to check if an external tool is available
function Test-Tool {
    param ([string]$tool)
    try { 
        Get-Command $tool -ErrorAction Stop | Out-Null 
        return $true 
    } catch { 
        return $false 
    }
}

# Function to handle file conversion with error handling
function Convert-File {
    param (
        [string]$filePath,
        [string]$outputFormat
    )
    $fileInfo = Get-Item $filePath
    $directory = $fileInfo.DirectoryName
    $baseName = $fileInfo.BaseName
    $extension = $fileInfo.Extension.ToLower()
    $outputFile = "$directory\$baseName.$outputFormat.ToLower()"

    # Check write permissions
    try {
        [System.IO.File]::WriteAllText("$directory\test.temp", "test") | Out-Null
        Remove-Item "$directory\test.temp" -Force
    } catch {
        Write-Host "│ Error: Insufficient permissions in $directory." -ForegroundColor Red
        return
    }

    Write-Host "│ Converting $filePath to $outputFormat..." -ForegroundColor White

    try {
        switch ($extension) {
            ".txt" {
                switch ($outputFormat) {
                    "CSV"  { Get-Content $filePath -ErrorAction Stop | ConvertTo-Csv -NoTypeInformation | Set-Content $outputFile -ErrorAction Stop }
                    "JSON" { Get-Content $filePath -ErrorAction Stop | ConvertTo-Json | Set-Content $outputFile -ErrorAction Stop }
                    "XML"  { [xml]$xml = "<root><text>$(Get-Content $filePath -ErrorAction Stop)</text></root>"; $xml.Save($outputFile) }
                    "HTML" { "<html><body><p>$(Get-Content $filePath -ErrorAction Stop)</p></body></html>" | Set-Content $outputFile -ErrorAction Stop }
                }
            }
            ".csv" {
                switch ($outputFormat) {
                    "JSON" { Import-Csv $filePath -ErrorAction Stop | ConvertTo-Json | Set-Content $outputFile -ErrorAction Stop }
                    "XML"  { Import-Csv $filePath -ErrorAction Stop | ConvertTo-Xml -As String | Set-Content $outputFile -ErrorAction Stop }
                    "XLSX" { 
                        if (-not (Get-Module -ListAvailable -Name ImportExcel)) { throw "ImportExcel module not installed." }
                        Import-Csv $filePath -ErrorAction Stop | Export-Excel -Path $outputFile -ErrorAction Stop 
                    }
                    "TXT"  { Import-Csv $filePath -ErrorAction Stop | ConvertTo-Csv -NoTypeInformation | Set-Content $outputFile -ErrorAction Stop }
                }
            }
            ".pdf" {
                switch ($outputFormat) {
                    "TXT"  { if (-not (Test-Tool "pdftotext")) { throw "pdftotext not installed." }; & pdftotext $filePath $outputFile }
                    "JPG"  { if (-not (Test-Tool "magick")) { throw "ImageMagick not installed." }; & magick convert $filePath $outputFile }
                    "PNG"  { if (-not (Test-Tool "magick")) { throw "ImageMagick not installed." }; & magick convert $filePath $outputFile }
                    "HTML" { if (-not (Test-Tool "pdftohtml")) { throw "pdftohtml not installed." }; & pdftohtml $filePath $outputFile }
                    "DOCX" { if (-not (Test-Tool "libreoffice")) { throw "LibreOffice not installed." }; & libreoffice --headless --convert-to docx $filePath --outdir $directory }
                }
            }
            default { Write-Host "│ Note: Conversion for $extension to $outputFormat not yet implemented." -ForegroundColor Yellow }
        }
        if (Test-Path $outputFile) {
            Write-Host "│ Success: File saved as $outputFile" -ForegroundColor Green
        } else {
            Write-Host "│ Error: Conversion failed - output file not created." -ForegroundColor Red
        }
    } catch {
        Write-Host "│ Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
}

# Main loop with improved UI
while ($true) {
    Show-Header
    $filePath = Read-Host "│ Enter file path: "
    
    if ($filePath -eq "exit") { break }
    if ($filePath -eq "help") { Show-Help; continue }
    
    $filePath = $filePath.Trim('"')
    
    if (-not (Test-Path $filePath)) {
        Write-Host "│ Error: File not found. Please try again." -ForegroundColor Red
        Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
        Start-Sleep -Seconds 2
        continue
    }

    $extension = (Get-Item $filePath).Extension.ToLower()
    if ($conversionOptions.ContainsKey($extension)) {
        Write-Host "├────────────────────────────────────────────────┤" -ForegroundColor Gray
        Write-Host "│ Available conversions for ${extension}:" -ForegroundColor White
        Write-Host "│" -ForegroundColor Gray
        $options = $conversionOptions[$extension]
        for ($i = 0; $i -lt $options.Count; $i++) {
            Write-Host "│   $($i + 1). $($options[$i])" -ForegroundColor White
        }
        Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
        
        $choice = Read-Host "│ Select an option (1-$($options.Count)): "
        if ($choice -eq "back") { continue }
        
        try {
            $choiceIndex = [int]$choice - 1
            if ($choiceIndex -ge 0 -and $choiceIndex -lt $options.Count) {
                Write-Host "├────────────────────────────────────────────────┤" -ForegroundColor Gray
                Convert-File -filePath $filePath -outputFormat $options[$choiceIndex]
                Read-Host "│ Press Enter to continue"
            } else {
                Write-Host "│ Error: Invalid option. Select 1-$($options.Count)." -ForegroundColor Red
                Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
                Start-Sleep -Seconds 2
            }
        } catch {
            Write-Host "│ Error: Please enter a valid number." -ForegroundColor Red
            Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
            Start-Sleep -Seconds 2
        }
    } else {
        Write-Host "│ Error: Unsupported file type." -ForegroundColor Red
        Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
        Start-Sleep -Seconds 2
    }
}

Write-Host "┌────────────────────────────────────────────────┐" -ForegroundColor Gray
Write-Host "│        Thank you for using FileFlex!           │" -ForegroundColor White
Write-Host "└────────────────────────────────────────────────┘" -ForegroundColor Gray
Read-Host "│ Press Enter to exit"