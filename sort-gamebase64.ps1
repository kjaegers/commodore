<#
.SYNOPSIS
    Organizes Commodore 64 GameBase archives by genre.

.DESCRIPTION
    Recursively scans for .ZIP files containing C64 games, extracts them,
    reads the .NFO file to determine the game's genre and name, then organizes
    the files into a genre/subgenre/gamename folder structure.
    
    Only processes games with English language support by default. Use -language option
    to select other languages.

    Extract your GameBase64 files and point this script (-inputPath) to the "Games" folder.
    I use a ramdisk for the extraction, temporary, and destination locations for speed purposes
    and to save on SSD wear.

.PARAMETER InputPath
    Path to scan for .ZIP files (recursively).

.PARAMETER OutputPath
    Path where organized games will be placed.

.PARAMETER TempPath
    Temporary extraction folder. Defaults to R:\work.

.PARAMETER Language
    Required language for games. Defaults to "English".

.EXAMPLE
    .\Sort-GameBase64.ps1 -InputPath "D:\C64Games" -OutputPath "E:\Organized"
    Organizes all C64 games from D:\C64Games into E:\Organized.

.EXAMPLE
    .\Sort-GameBase64.ps1 -InputPath "D:\C64Games" -OutputPath "E:\Organized" -TempPath "C:\Temp"
    Uses C:\Temp as the temporary extraction folder.

.EXAMPLE
    .\Sort-GameBase64.ps1 -InputPath "D:\C64Games" -OutputPath "E:\Organized" -Language "German"
    Organizes only German language games.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,
    
    [Parameter(Mandatory = $true, Position = 1)]
    [string]$OutputPath,
    
    [Parameter()]
    [string]$TempPath = "R:\work",
    
    [Parameter()]
    [string]$Language = "English"
)

# Validate paths
if (-not (Test-Path $InputPath)) {
    Write-Error "Input path does not exist: $InputPath"
    exit 1
}

if (-not (Test-Path $OutputPath)) {
    try {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
    } catch {
        Write-Error "Failed to create output directory: $($_.Exception.Message)"
        exit 1
    }
}

if (-not (Test-Path $TempPath)) {
    try {
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
        Write-Host "Created temp directory: $TempPath" -ForegroundColor Green
    } catch {
        Write-Error "Failed to create temp directory: $($_.Exception.Message)"
        exit 1
    }
}

# Function to clean invalid characters from folder/file names
function Clean-FileName {
    param([string]$Name)
    
    # Remove square brackets
    $cleaned = $Name -replace '[\[\]]', ''
    
    # Remove other invalid filename characters
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalidChars) {
        $cleaned = $cleaned -replace [regex]::Escape($char), ''
    }
    
    # Remove leading/trailing spaces and dots
    $cleaned = $cleaned.Trim('. ')
    
    # Replace multiple spaces with single space
    $cleaned = $cleaned -replace '\s+', ' '
    
    return $cleaned
}

# Function to extract ZIP file
function Expand-ZipToTemp {
    param(
        [string]$ZipPath,
        [string]$TempFolder
    )
    
    try {
        # Create unique temp subfolder for this extraction
        $tempSubFolder = Join-Path $TempFolder ([System.IO.Path]::GetFileNameWithoutExtension($ZipPath))
        if (Test-Path $tempSubFolder) {
            Remove-Item $tempSubFolder -Recurse -Force -Verbose:$false | Out-Null
        }
        New-Item -ItemType Directory -Path $tempSubFolder -Force | Out-Null
        
        # Extract using .NET
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $tempSubFolder)
        
        return $tempSubFolder
    } catch {
        Write-Warning "Failed to extract $ZipPath`: $($_.Exception.Message)"
        return $null
    }
}

# Function to parse NFO file
function Get-GameInfo {
    param([string]$NfoPath)
    
    if (-not (Test-Path $NfoPath)) {
        return $null
    }
    
    try {
        $content = Get-Content $NfoPath -Raw
        
        $gameInfo = @{
            Name = $null
            Genre = $null
            SubGenre = $null
            Language = $null
        }
        
        # Extract Name
        if ($content -match '(?m)^Name:\s+(.+)$') {
            $gameInfo.Name = $matches[1].Trim()
        }
        
        # Extract Language
        if ($content -match '(?m)^Language:\s+(.+)$') {
            $gameInfo.Language = $matches[1].Trim()
        }
        
        # Extract Genre (may include subgenre after " - ")
        if ($content -match '(?m)^Genre:\s+(.+)$') {
            $genreText = $matches[1].Trim()
            
            # Check for subgenre separator
            if ($genreText -match '^(.+?)\s+-\s+(.+)$') {
                $gameInfo.Genre = $matches[1].Trim()
                $gameInfo.SubGenre = $matches[2].Trim()
            } else {
                $gameInfo.Genre = $genreText
                $gameInfo.SubGenre = $null
            }
        }
        
        return $gameInfo
    } catch {
        Write-Warning "Failed to parse NFO file: $($_.Exception.Message)"
        return $null
    }
}

# Suppress progress messages from Remove-Item
$ProgressPreference = 'SilentlyContinue'

# Main processing
Write-Host "`nCommodore 64 Game Organizer" -ForegroundColor Cyan
Write-Host "==========================`n" -ForegroundColor Cyan
Write-Host "Input Path:  $InputPath" -ForegroundColor White
Write-Host "Output Path: $OutputPath" -ForegroundColor White
Write-Host "Temp Path:   $TempPath" -ForegroundColor White
Write-Host "Language:    $Language`n" -ForegroundColor White

# Find all ZIP files
Write-Host "Scanning for ZIP files..." -ForegroundColor Yellow
$zipFiles = Get-ChildItem -Path $InputPath -Filter "*.zip" -Recurse -File | Sort-Object FullName

if ($zipFiles.Count -eq 0) {
    Write-Host "No ZIP files found." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($zipFiles.Count) ZIP files to process`n" -ForegroundColor Green

# Statistics
$stats = @{
    Total = $zipFiles.Count
    Processed = 0
    Skipped = 0
    SkippedWrongLanguage = 0
    SkippedNoNFO = 0
    SkippedNoInfo = 0
    Errors = 0
}

$currentIndex = 0

foreach ($zipFile in $zipFiles) {
    $currentIndex++
    
    Write-Host "[$currentIndex/$($stats.Total)] Processing: $($zipFile.Name)" -ForegroundColor Cyan
    
    # Extract to temp folder
    $tempFolder = Expand-ZipToTemp -ZipPath $zipFile.FullName -TempFolder $TempPath
    
    if (-not $tempFolder) {
        Write-Host "  ✗ Failed to extract" -ForegroundColor Red
        $stats.Errors++
        continue
    }
    
    try {
        # Find NFO file
        $nfoFile = Get-ChildItem -Path $tempFolder -Filter "*.nfo" -File | Select-Object -First 1
        
        if (-not $nfoFile) {
            Write-Host "  ⊘ Skipping - No NFO file found" -ForegroundColor Yellow
            $stats.Skipped++
            $stats.SkippedNoNFO++
            continue
        }
        
        Write-Host "  → Found NFO: $($nfoFile.Name)" -ForegroundColor Gray
        
        # Parse NFO
        $gameInfo = Get-GameInfo -NfoPath $nfoFile.FullName
        
        if (-not $gameInfo -or -not $gameInfo.Name -or -not $gameInfo.Genre) {
            Write-Host "  ⊘ Skipping - Could not extract game info from NFO" -ForegroundColor Yellow
            $stats.Skipped++
            $stats.SkippedNoInfo++
            continue
        }
        
        Write-Host "  → Name: $($gameInfo.Name)" -ForegroundColor Gray
        Write-Host "  → Genre: $($gameInfo.Genre)" -ForegroundColor Gray
        if ($gameInfo.SubGenre) {
            Write-Host "  → SubGenre: $($gameInfo.SubGenre)" -ForegroundColor Gray
        }
        Write-Host "  → Language: $($gameInfo.Language)" -ForegroundColor Gray
        
        # Check language
        if ($gameInfo.Language -notmatch $Language) {
            Write-Host "  ⊘ Skipping - Language is '$($gameInfo.Language)' (required: $Language)" -ForegroundColor Yellow
            $stats.Skipped++
            $stats.SkippedWrongLanguage++
            continue
        }
        
        # Clean names
        $cleanGenre = Clean-FileName $gameInfo.Genre
        $cleanSubGenre = if ($gameInfo.SubGenre) { Clean-FileName $gameInfo.SubGenre } else { $null }
        $cleanGameName = Clean-FileName $gameInfo.Name
        
        if ([string]::IsNullOrWhiteSpace($cleanGenre) -or [string]::IsNullOrWhiteSpace($cleanGameName)) {
            Write-Host "  ⊘ Skipping - Invalid genre or game name after cleaning" -ForegroundColor Yellow
            $stats.Skipped++
            continue
        }
        
        # Build destination path
        $destPath = Join-Path $OutputPath $cleanGenre
        
        if ($cleanSubGenre) {
            $destPath = Join-Path $destPath $cleanSubGenre
        }
        
        $destPath = Join-Path $destPath $cleanGameName
        
        # Create destination folders
        if (-not (Test-Path $destPath)) {
            New-Item -ItemType Directory -Path $destPath -Force | Out-Null
            Write-Host "  → Created folder: $($destPath.Substring($OutputPath.Length))" -ForegroundColor Green
        }
        
        # Move files from temp to destination
        $filesInTemp = Get-ChildItem -Path $tempFolder -File
        $movedCount = 0
        
        foreach ($file in $filesInTemp) {
            $destFile = Join-Path $destPath $file.Name
            
            # Handle duplicates
            if (Test-Path $destFile) {
                $counter = 1
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $extension = [System.IO.Path]::GetExtension($file.Name)
                
                while (Test-Path $destFile) {
                    $newName = "${baseName}_${counter}${extension}"
                    $destFile = Join-Path $destPath $newName
                    $counter++
                }
                
                Write-Host "  → File exists, using: $(Split-Path $destFile -Leaf)" -ForegroundColor DarkYellow
            }
            
            Move-Item -LiteralPath $file.FullName -Destination $destFile -Force
            $movedCount++
        }
        
        Write-Host "  ✓ Moved $movedCount files to: $($destPath.Substring($OutputPath.Length))" -ForegroundColor Green
        $stats.Processed++
        
    } catch {
        Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        $stats.Errors++
    } finally {
        # Clean up temp folder
        if (Test-Path $tempFolder) {
            Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue -Verbose:$false | Out-Null
        }
    }
    
    Write-Host ""
}

# Display summary
Write-Host "=================================="
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "==================================`n"
Write-Host "Total ZIP files:        $($stats.Total)"
Write-Host "Successfully processed: $($stats.Processed) " -NoNewline
Write-Host "✓" -ForegroundColor Green
Write-Host "Skipped (total):        $($stats.Skipped) " -NoNewline
if ($stats.Skipped -gt 0) {
    Write-Host "⊘" -ForegroundColor Yellow
} else {
    Write-Host ""
}
Write-Host "  - Wrong language:     $($stats.SkippedWrongLanguage)"
Write-Host "  - No NFO file:        $($stats.SkippedNoNFO)"
Write-Host "  - Missing info:       $($stats.SkippedNoInfo)"
Write-Host "Errors:                 $($stats.Errors) " -NoNewline
if ($stats.Errors -gt 0) {
    Write-Host "✗" -ForegroundColor Red
} else {
    Write-Host ""
}

Write-Host "`nOrganization complete!" -ForegroundColor Green
Write-Host ""

exit 0
