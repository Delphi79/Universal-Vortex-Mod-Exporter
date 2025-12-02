[CmdletBinding()]
param()

# ---------- Helpers ----------

function Format-GameLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    # replace underscores with spaces
    $t = $Name -replace '_', ' '

    # insert space before trailing digits (fallout4 -> fallout 4)
    if ($t -match '^(.*?)(\d+)$') {
        $t = "$($matches[1]) $($matches[2])"
    }

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $label   = $culture.TextInfo.ToTitleCase($t.ToLower())

    return $label.Trim()
}

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

            # Load order (index or n/a)
            $loadOrder = "n/a"
            if ($deployMap.Count -gt 0 -and $deployMap.ContainsKey($m.id)) {
                $loadOrder = [string]$deployMap[$m.id]
            }

            $mods += [PSCustomObject]@{
                GameName   = $gameName
                ModId      = $m.id
                ModName    = $m.attributes.modName
                ModVersion = $m.attributes.modVersion
                Enabled    = $enabled
                LoadOrder  = $loadOrder
                Source     = $m.attributes.source
                Homepage   = $m.attributes.homepage
            }
        }
    }

    return $mods
}

# ---------- Main ----------

Write-Host ""
Write-Host "Universal Vortex Mod Exporter (UVME)" -ForegroundColor Cyan
Write-Host "-----------------------------------" -ForegroundColor Cyan

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

# Build distinct game list + nice labels
$gamesRaw = $allMods | Select-Object GameName -Unique | Sort-Object GameName

$games = @()
$idx   = 1
foreach ($g in $gamesRaw) {
    $games += [PSCustomObject]@{
        Index       = $idx
        GameName    = $g.GameName
        DisplayName = Format-GameLabel $g.GameName
    }
    $idx++
}

Write-Host ""
Write-Host "Games found in Vortex backup:" -ForegroundColor Yellow
foreach ($g in $games) {
    Write-Host ("{0}) {1}" -f $g.Index, $g.DisplayName)
}

Write-Host ""
$gameChoice = Read-Host "Enter a game number to export only that game, or press Enter for ALL games"

if ([string]::IsNullOrWhiteSpace($gameChoice)) {
    $modsForScope = $allMods
    $gameLabel    = "AllGames"
} else {
    if (-not ($gameChoice -as [int])) {
        Write-Error "Invalid selection. Please enter a number next time."
        exit 1
    }
    $choiceInt = [int]$gameChoice
    $selected  = $games | Where-Object { $_.Index -eq $choiceInt }

    if (-not $selected) {
        Write-Error "No game with that number. Aborting."
        exit 1
    }

    $modsForScope = $allMods | Where-Object { $_.GameName -eq $selected.GameName }
    $gameLabel    = $selected.DisplayName

    if (-not $modsForScope -or $modsForScope.Count -eq 0) {
        Write-Error "No mods found for game '$($selected.DisplayName)'."
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
Write-Host "4) Export as HTML (table with clickable links)"
$formatChoice = Read-Host "Choose 1, 2, 3 or 4"

$basePath = $PSScriptRoot
if (-not $basePath) { $basePath = (Get-Location).Path }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileBase  = "${gameLabel}_${scopeLabel}_${timestamp}"

# Common sorted & projected data
$sortedMods  = $modsToExport | Sort-Object GameName, LoadOrder, ModName
$selectProps = "GameName","ModName","ModVersion","Enabled","LoadOrder","Source","Homepage"

# Layout helper for Excel
function Set-NormalLandscapeLayout {
    param($Sheet)

    try {
        $ps = $Sheet.PageSetup
        # 2 = xlLandscape
        $ps.Orientation = 2
        $ps.Zoom        = 100
    } catch { }

    try {
        $Sheet.Columns.Item(1).ColumnWidth = 12   # GameName
        $Sheet.Columns.Item(2).ColumnWidth = 45   # ModName
        $Sheet.Columns.Item(3).ColumnWidth = 10   # ModVersion
        $Sheet.Columns.Item(4).ColumnWidth = 9    # Enabled
        $Sheet.Columns.Item(5).ColumnWidth = 14   # LoadOrder
        $Sheet.Columns.Item(6).ColumnWidth = 10   # Source
        $Sheet.Columns.Item(7).ColumnWidth = 12   # Homepage
        $Sheet.Columns.Item(2).WrapText    = $true
    } catch { }
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
            $sheet.Cells.Item($row, 5).Value2 = [string]$m.LoadOrder
            $sheet.Cells.Item($row, 6).Value2 = [string]$m.Source

            $url = $m.Homepage
            if ($url -and $url -is [string] -and $url.Trim().StartsWith("http")) {
                $cell = $sheet.Cells.Item($row, 7)
                $cell.Value2 = "Download"
                $sheet.Hyperlinks.Add($cell, $url, "", "", "Download") | Out-Null
            } else {
                $sheet.Cells.Item($row, 7).Value2 = [string]$url
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
        # HTML (with clickable links + banner + avatar)
        $outFile = Join-Path $basePath "$fileBase.html"

        $rows = foreach ($m in $sortedMods) {
            $gameName   = [System.Net.WebUtility]::HtmlEncode([string]$m.GameName)
            $modName    = [System.Net.WebUtility]::HtmlEncode([string]$m.ModName)
            $modVersion = [System.Net.WebUtility]::HtmlEncode([string]$m.ModVersion)
            $enabled    = [System.Net.WebUtility]::HtmlEncode([string]$m.Enabled)
            $loadOrder  = [System.Net.WebUtility]::HtmlEncode([string]$m.LoadOrder)
            $source     = [System.Net.WebUtility]::HtmlEncode([string]$m.Source)

            $url = [string]$m.Homepage
            if ($url -and $url.Trim().StartsWith("http")) {
                $urlEsc = [System.Net.WebUtility]::HtmlEncode($url)
                $homepageCell = "<a href=""$urlEsc"" target=""_blank"" rel=""noopener"">Download</a>"
            } else {
                $homepageCell = [System.Net.WebUtility]::HtmlEncode($url)
            }

            "<tr>
                <td>$gameName</td>
                <td>$modName</td>
                <td>$modVersion</td>
                <td>$enabled</td>
                <td>$loadOrder</td>
                <td>$source</td>
                <td>$homepageCell</td>
            </tr>"
        }

        $title     = "Universal Vortex Mod Exporter - $gameLabel ($scopeLabel)"
        $bannerImg = "images/uvme-banner.png"
        $avatarImg = "images/vaultboy.png"

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <title>$title</title>
    <style>
        body {
            font-family: Segoe UI, Arial, sans-serif;
            margin: 20px;
            background-color: #18181b;
            color: #f4f4f5;
        }
        h1 {
            font-size: 20px;
            margin-bottom: 4px;
        }
        h2 {
            font-size: 14px;
            font-weight: normal;
            color: #a1a1aa;
            margin-top: 0;
        }
        .meta {
            font-size: 11px;
            color: #a1a1aa;
            margin-bottom: 4px;
        }

        .header-visual {
            display: flex;
            justify-content: center;
            align-items: flex-end;
            gap: 16px;
            margin-bottom: 8px;
        }
        .banner-wrap {
            flex: 1;
            height: 120px;
            overflow: hidden;
            display: flex;
            justify-content: center;
            align-items: flex-end;
        }
        .banner-wrap img {
            height: 120px;
            width: auto;
            border-radius: 6px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.6);
        }
        .avatar-wrap {
            flex: 0 0 auto;
            padding-bottom: 4px;
            padding-right: 4px;
        }
        .avatar-wrap img {
            width: 56px;
            height: 56px;
            border-radius: 50%;
            box-shadow:
                0 0 0 2px #f97316,
                0 0 10px rgba(249, 115, 22, 0.8);
        }

        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 12px;
            font-size: 13px;
        }
        th, td {
            border: 1px solid #3f3f46;
            padding: 6px 8px;
            vertical-align: top;
        }
        th {
            background-color: #27272a;
            position: sticky;
            top: 0;
            z-index: 1;
        }
        tr:nth-child(even) {
            background-color: #18181b;
        }
        tr:nth-child(odd) {
            background-color: #09090b;
        }
        a {
            color: #22c55e;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>

    <div class="header-visual">
        <div class="banner-wrap">
            <img src="$bannerImg" alt="Universal Vortex Mod Exporter" />
        </div>
        <div class="avatar-wrap">
            <img src="$avatarImg" alt="Vault Boy" />
        </div>
    </div>

    <h1>Universal Vortex Mod Exporter</h1>
    <h2>$gameLabel &mdash; $scopeLabel</h2>

    <div class="meta">
        Generated: $(Get-Date)<br />
        Total mods: $($sortedMods.Count)
    </div>

    <table>
        <thead>
            <tr>
                <th>Game</th>
                <th>Mod Name</th>
                <th>Version</th>
                <th>Enabled</th>
                <th>Load Order</th>
                <th>Source</th>
                <th>Homepage</th>
            </tr>
        </thead>
        <tbody>
            $($rows -join "`r`n")
        </tbody>
    </table>
</body>
</html>
"@

        $html | Set-Content -Path $outFile -Encoding UTF8
        Write-Host "`nHTML exported to (open this in your browser):`n  $outFile" -ForegroundColor Green
    }
    default {
        Write-Error "Invalid choice. Aborting."
        exit 1
    }
}

Write-Host "`nDone." -ForegroundColor Cyan
