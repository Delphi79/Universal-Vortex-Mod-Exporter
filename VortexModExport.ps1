[CmdletBinding()]
param()

# Read latest full backup JSON Vortex keeps in its temp folder
function Get-VortexSnapshot {
    $pattern = Join-Path $env:APPDATA "Vortex\temp\state_backups_full\*.json"

    $latest = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

    if (-not $latest) {
        throw "No Vortex backup JSON found at $pattern. Start Vortex once and try again."
    }

    Get-Content -Path $latest.FullName -Raw | ConvertFrom-Json
}

# Turn snapshot into a flat list of mods with optional deployment order
function Get-VortexModOverview {
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot
    )

    $mods = @()

    # gameName → gameId mapping
    $profileMap = $Snapshot.settings.profiles.lastActiveProfile
    if (-not $profileMap) {
        throw "No lastActiveProfile section in Vortex backup."
    }

    foreach ($prop in $profileMap.PSObject.Properties) {
        $gameName = $prop.Name
        $gameId   = $prop.Value

        $modBlock = $Snapshot.persistent.mods.$gameName
        if (-not $modBlock) { continue }

        $profileState = $Snapshot.persistent.profiles.$gameId.modState

        # Build deployment map (modId -> index) if this game exposes it
        $deployMap = @{}
        $loadOrderArray = $Snapshot.persistent.loadOrder.$gameId
        if ($loadOrderArray) {
            for ($i = 0; $i -lt $loadOrderArray.Count; $i++) {
                $entry = $loadOrderArray[$i]
                if ($entry -and $entry.modId) {
                    $deployMap[$entry.modId] = $i
                }
            }
        }

        foreach ($modKey in $modBlock.PSObject.Properties.Name) {
            $m = $modBlock.$modKey

            # Enabled flag from profile
            $enabled = $false
            if ($profileState -and $profileState.$modKey) {
                $enabled = [bool]$profileState.$modKey.enabled
            }

            # Deployment order + flag
            $deploymentOrder     = "n/a"
            $loadOrderFromVortex = "No"

            if ($deployMap.Count -gt 0 -and $deployMap.ContainsKey($m.id)) {
                $deploymentOrder     = [string]$deployMap[$m.id]
                $loadOrderFromVortex = "Yes"
            }

            $mods += [PSCustomObject]@{
                GameName            = $gameName
                ModId               = $m.id
                ModName             = $m.attributes.modName
                ModVersion          = $m.attributes.modVersion
                Enabled             = $enabled
                DeploymentOrder     = $deploymentOrder
                LoadOrderFromVortex = $loadOrderFromVortex
                Source              = $m.attributes.source
                Homepage            = $m.attributes.homepage
            }
        }
    }

    return $mods
}

Write-Host "   = Vortex Mod Exporter =" -ForegroundColor Cyan
Write-Host "     ==== ImaNewb79 ====" -ForegroundColor Cyan
Write-Host "      =================" -ForegroundColor Cyan
Write-Host "       ===============" -ForegroundColor Cyan
Write-Host "        =============" -ForegroundColor Cyan
Write-Host "         ===========" -ForegroundColor Cyan
Write-Host "          =========" -ForegroundColor Cyan
Write-Host "           =======" -ForegroundColor Cyan
Write-Host "            =====" -ForegroundColor Cyan
Write-Host "            ===" -ForegroundColor Cyan
Write-Host "            ==" -ForegroundColor Cyan
Write-Host "           =" -ForegroundColor Cyan
Write-Host "          =" -ForegroundColor Cyan



try {
    $snapshot = Get-VortexSnapshot
} catch {
    Write-Error $_
    exit 1
}

try {
    $allMods = Get-VortexModOverview -Snapshot $snapshot
} catch {
    Write-Error $_
    exit 1
}

if (-not $allMods -or $allMods.Count -eq 0) {
    Write-Error "No mods found in Vortex backup."
    exit 1
}

# Show discoveCyan games
$games = $allMods | Select-Object GameName -Unique
Write-Host ""
Write-Host "Games found in Vortex backup:" -ForegroundColor Yellow
$games | ForEach-Object { Write-Host " - $($_.GameName)" }

# Optional: filter by game
Write-Host ""
$gameFilter = Read-Host "Enter a game name to export only that game, or press Enter for ALL games"

if ([string]::IsNullOrWhiteSpace($gameFilter)) {
    $modsForScope = $allMods
    $gameLabel    = "AllGames"
} else {
    $modsForScope = $allMods | Where-Object { $_.GameName -eq $gameFilter }
    $gameLabel    = $gameFilter

    if (-not $modsForScope -or $modsForScope.Count -eq 0) {
        Write-Error "No mods found for game '$gameFilter'. Check the name (exact match) and try again."
        exit 1
    }
}

# All vs enabled-only
Write-Host ""
Write-Host "1) Export ALL mods"
Write-Host "2) Export ONLY ENABLED mods"
$scopeChoice = Read-Host "Choose 1 or 2"

switch ($scopeChoice) {
    "1" {
        $modsToExport = $modsForScope
        $scopeLabel   = "AllMods"
    }
    "2" {
        $modsToExport = $modsForScope | Where-Object { $_.Enabled -eq $true }
        $scopeLabel   = "EnabledMods"
    }
    default {
        Write-Error "Invalid choice. Aborting."
        exit 1
    }
}

if (-not $modsToExport -or $modsToExport.Count -eq 0) {
    Write-Error "No mods matching selection (maybe no enabled mods?). Aborting."
    exit 1
}

# Output format selection
Write-Host ""
Write-Host "1) Export as CSV"
Write-Host "2) Export as JSON"
Write-Host "3) Export as Excel (XLSX with clickable links)"
Write-Host "4) Export as PDF (with clickable links)"
$formatChoice = Read-Host "Choose 1, 2, 3 or 4"

$basePath = $PSScriptRoot
if (-not $basePath) { $basePath = (Get-Location).Path }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileBase  = "${gameLabel}_${scopeLabel}_${timestamp}"

# Common sorted & projected data
$sortedMods  = $modsToExport | Sort-Object GameName, DeploymentOrder, ModName
$selectProps = "GameName","ModName","ModVersion","Enabled","DeploymentOrder","LoadOrderFromVortex","Source","Homepage"

# Helper: set sane layout for printing / PDF
function Set-NormalLandscapeLayout {
    param($Sheet)

    try {
        $ps = $Sheet.PageSetup
        # 2 = xlLandscape
        $ps.Orientation = 2
        $ps.Zoom        = 100   # no auto-shrink; normal readable size
    } catch {
        # If Excel is grumpy, ignore – worst case it uses defaults.
    }

    # Fixed column widths so Excel doesn't go wild
    try {
        $Sheet.Columns.Item(1).ColumnWidth = 12   # GameName
        $Sheet.Columns.Item(2).ColumnWidth = 45   # ModName (wrapped)
        $Sheet.Columns.Item(3).ColumnWidth = 10   # ModVersion
        $Sheet.Columns.Item(4).ColumnWidth = 9    # Enabled
        $Sheet.Columns.Item(5).ColumnWidth = 14   # DeploymentOrder
        $Sheet.Columns.Item(6).ColumnWidth = 18   # LoadOrderFromVortex
        $Sheet.Columns.Item(7).ColumnWidth = 10   # Source
        $Sheet.Columns.Item(8).ColumnWidth = 12   # Homepage (Download)

        # Wrap long mod names in the ModName column
        $Sheet.Columns.Item(2).WrapText = $true
    } catch {
        # again, not fatal if it fails
    }
}

switch ($formatChoice) {
    "1" {
        # CSV
        $outFile = Join-Path $basePath "$fileBase.csv"
        $sortedMods |
            Select-Object $selectProps |
            Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

        Write-Host "`nCSV exported to:`n  $outFile" -ForegroundColor Green
    }
    "2" {
        # JSON
        $outFile = Join-Path $basePath "$fileBase.json"
        $sortedMods |
            Select-Object $selectProps |
            ConvertTo-Json -Depth 5 |
            Out-File -FilePath $outFile -Encoding UTF8

        Write-Host "`nJSON exported to:`n  $outFile" -ForegroundColor Green
    }
    "3" {
        # Excel (XLSX) with clickable links
        $outFile = Join-Path $basePath "$fileBase.xlsx"

        $excel         = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook      = $excel.Workbooks.Add()
        $sheet         = $workbook.Worksheets.Item(1)

        $headers = $selectProps

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        }

        $row = 2
        foreach ($m in $sortedMods) {
            $sheet.Cells.Item($row, 1).Value2 = [string]$m.GameName
            $sheet.Cells.Item($row, 2).Value2 = [string]$m.ModName
            $sheet.Cells.Item($row, 3).Value2 = [string]$m.ModVersion
            $sheet.Cells.Item($row, 4).Value2 = [string]$m.Enabled
            $sheet.Cells.Item($row, 5).Value2 = [string]$m.DeploymentOrder
            $sheet.Cells.Item($row, 6).Value2 = [string]$m.LoadOrderFromVortex
            $sheet.Cells.Item($row, 7).Value2 = [string]$m.Source

            $url = $m.Homepage
            if ($url -and $url -is [string] -and $url.Trim().StartsWith("http")) {
                $cell = $sheet.Cells.Item($row, 8)
                $cell.Value2 = "Download"
                $sheet.Hyperlinks.Add($cell, $url, "", "", "Download") | Out-Null
            } else {
                $sheet.Cells.Item($row, 8).Value2 = [string]$url
            }

            $row++
        }

        Set-NormalLandscapeLayout -Sheet $sheet

        $workbook.SaveAs($outFile)
        $workbook.Close($true)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null

        Write-Host "`nExcel file exported to:`n  $outFile" -ForegroundColor Green
    }
    "4" {
        # PDF (via Excel, preserving hyperlinks)
        $outFile = Join-Path $basePath "$fileBase.pdf"

        $excel         = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook      = $excel.Workbooks.Add()
        $sheet         = $workbook.Worksheets.Item(1)

        $headers = $selectProps

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        }

        $row = 2
        foreach ($m in $sortedMods) {
            $sheet.Cells.Item($row, 1).Value2 = [string]$m.GameName
            $sheet.Cells.Item($row, 2).Value2 = [string]$m.ModName
            $sheet.Cells.Item($row, 3).Value2 = [string]$m.ModVersion
            $sheet.Cells.Item($row, 4).Value2 = [string]$m.Enabled
            $sheet.Cells.Item($row, 5).Value2 = [string]$m.DeploymentOrder
            $sheet.Cells.Item($row, 6).Value2 = [string]$m.LoadOrderFromVortex
            $sheet.Cells.Item($row, 7).Value2 = [string]$m.Source

            $url = $m.Homepage
            if ($url -and $url -is [string] -and $url.Trim().StartsWith("http")) {
                $cell = $sheet.Cells.Item($row, 8)
                $cell.Value2 = "Download"
                $sheet.Hyperlinks.Add($cell, $url, "", "", "Download") | Out-Null
            } else {
                $sheet.Cells.Item($row, 8).Value2 = [string]$url
            }

            $row++
        }

        Set-NormalLandscapeLayout -Sheet $sheet

        # 0 = PDF for ExportAsFixedFormat
        $xlFixedFormatType_PDF = 0
        $workbook.ExportAsFixedFormat($xlFixedFormatType_PDF, $outFile)

        $workbook.Close($false)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null

        Write-Host "`nPDF exported to:`n  $outFile" -ForegroundColor Green
    }
    default {
        Write-Error "Invalid choice. Aborting."
        exit 1
    }
}

Write-Host "`nDone." -ForegroundColor Cyan
