[CmdletBinding()]
param()

# ==========================================================
#  Helpers
# ==========================================================

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

function Get-VortexSnapshot {
    # Read latest full backup JSON Vortex keeps in its temp folder
    $pattern = Join-Path $env:APPDATA "Vortex\temp\state_backups_full\*.json"

    $latest = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

    if (-not $latest) {
        throw "No Vortex backup JSON found at $pattern. Start/restart Vortex so it writes a full state backup, then run UVME again."
    }

    Get-Content -Path $latest.FullName -Raw | ConvertFrom-Json
}

function Get-UVMEGameOverview {
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot
    )

    $modsRoot      = $Snapshot.persistent.mods
    $profilesRoot  = $Snapshot.persistent.profiles
    $lastActiveMap = $Snapshot.settings.profiles.lastActiveProfile

    if (-not $modsRoot) {
        throw "Snapshot has no persistent.mods section."
    }

    $rows = @()

    foreach ($gameProp in $modsRoot.PSObject.Properties) {
        $gameName  = $gameProp.Name       # e.g. "fallout4", "balatro"
        $modBlock  = $gameProp.Value
        if (-not $modBlock) { continue }

        # Try to find the profile used for this game
        $profileId   = $null
        if ($lastActiveMap -and $lastActiveMap.PSObject.Properties.Name -contains $gameName) {
            $profileId = $lastActiveMap.$gameName
        }

        $profileState = $null
        if ($profileId -and $profilesRoot -and $profilesRoot.PSObject.Properties.Name -contains $profileId) {
            $profileState = $profilesRoot.$profileId.modState
        }

        # IMPORTANT: deployment order = index of the mod within persistent.mods.<gameName>
        # Enumerate properties in the order they appear in JSON & use that index.
        $deployIndex = 0
        foreach ($modProp in $modBlock.PSObject.Properties) {
            $modKey = $modProp.Name
            $m      = $modProp.Value
            if (-not $m) { 
                $deployIndex++
                continue 
            }

            $attrs = $m.attributes

            # Enabled flag from (game) profile, if present
            $enabled = $false
            if ($profileState -and $profileState.PSObject.Properties.Name -contains $modKey) {
                $enabled = [bool]$profileState.$modKey.enabled
            }

            # Only persist Source/Homepage as strings; ignore booleans etc
            $source   = $null
            $homepage = $null
            if ($attrs) {
                if ($attrs.source -is [string] -and $attrs.source.Trim().Length -gt 0) {
                    $source = $attrs.source
                }
                if ($attrs.homepage -is [string] -and $attrs.homepage.Trim().Length -gt 0) {
                    $homepage = $attrs.homepage
                }
            }

            $rows += [PSCustomObject]@{
                GameName   = $gameName
                ModId      = $m.id
                ModKey     = $modKey
                ModName    = $attrs.modName
                ModVersion = $attrs.modVersion
                Enabled    = $enabled
                # "LoadOrder" column in the output is actually the Vortex deploy order
                LoadOrder  = [string]$deployIndex
                Source     = $source
                Homepage   = $homepage
            }

            $deployIndex++
        }
    }

    return $rows
}

function Get-UVMEArchiveOverview {
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot,

        [string]$GameFilter
    )

    $downloadsNode = $Snapshot.persistent.downloads
    if (-not $downloadsNode -or -not $downloadsNode.files) {
        return @()
    }

    # Base downloads dir is under %APPDATA%\Vortex\downloads\<game>
    $downloadsRoot = Join-Path $env:APPDATA "Vortex\downloads"

    $rows = @()

    foreach ($fileProp in $downloadsNode.files.PSObject.Properties) {
        $entry = $fileProp.Value
        if (-not $entry) { continue }

        $gamesForFile = $entry.game
        if (-not $gamesForFile) { continue }

        foreach ($g in $gamesForFile) {
            if ($GameFilter -and $g -ne $GameFilter) { continue }

            $gameName  = $g
            $fileName  = $entry.localPath
            $modName   = $null
            if ($entry.modInfo -and $entry.modInfo.name) {
                $modName = $entry.modInfo.name
            }

            $sizeBytes = [double]($entry.size   | ForEach-Object {$_})
            $sizeMB    = if ($sizeBytes -gt 0) { [math]::Round($sizeBytes / 1MB, 2) } else { 0 }

            $modified = $null
            if ($entry.fileTime) {
                try {
                    $modified = [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$entry.fileTime).LocalDateTime
                } catch {
                    $modified = $null
                }
            }

            $fullPath = $null
            $exists   = $false
            if ($fileName) {
                $gameFolder = $gameName
                $fullPath   = Join-Path (Join-Path $downloadsRoot $gameFolder) $fileName
                $exists     = Test-Path -LiteralPath $fullPath
            }

            $rows += [PSCustomObject]@{
                GameName    = $gameName
                FileName    = $fileName
                ModName     = $modName
                SizeMB      = $sizeMB
                Modified    = $modified
                ExistsOnDisk= $exists
                FullPath    = $fullPath
            }
        }
    }

    return $rows
}

function New-ExcelLayout {
    param(
        [Parameter(Mandatory = $true)]
        $Sheet
    )

    try {
        $ps = $Sheet.PageSetup
        # 2 = xlLandscape
        $ps.Orientation = 2
        $ps.Zoom        = 100
    } catch { }

    try {
        $Sheet.Columns.Item(1).ColumnWidth = 12   # Game
        $Sheet.Columns.Item(2).ColumnWidth = 45   # Mod / File Name
        $Sheet.Columns.Item(3).ColumnWidth = 18   # Version / Mod Name
        $Sheet.Columns.Item(4).ColumnWidth = 10   # Enabled / Size
        $Sheet.Columns.Item(5).ColumnWidth = 12   # Load Order / Modified
        $Sheet.Columns.Item(6).ColumnWidth = 12   # Source / Exists
        $Sheet.Columns.Item(7).ColumnWidth = 70   # Homepage / Full Path
        $Sheet.Columns.Item(2).WrapText    = $true
    } catch { }
}

function Test-UVMEIsUrl {
    param(
        [AllowNull()]
        $Value
    )

    if (-not $Value) { return $false }
    if (-not ($Value -is [string])) { return $false }

    $trim = $Value.Trim()
    return ($trim.StartsWith("http://") -or $trim.StartsWith("https://") -or $trim.StartsWith("ftp://"))
}

# ==========================================================
#  Main
# ==========================================================

Write-Host ""
Write-Host "Universal Vortex Mod Exporter (UVME)" -ForegroundColor Cyan
Write-Host "-----------------------------------" -ForegroundColor Cyan
Write-Host ""
Write-Host "TIP:" -ForegroundColor Yellow
Write-Host "  If you've recently changed mods/games in Vortex," -ForegroundColor Yellow
Write-Host "  restart Vortex so it writes a fresh backup before running UVME." -ForegroundColor Yellow
Write-Host ""

try {
    $snapshot = Get-VortexSnapshot
} catch {
    Write-Error $_
    Read-Host "Press Enter to exit"
    exit 1
}

# Build full mods table once
try {
    $allMods = Get-UVMEGameOverview -Snapshot $snapshot
} catch {
    Write-Error $_
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not $allMods -or $allMods.Count -eq 0) {
    Write-Error "No mods found in Vortex snapshots."
    Read-Host "Press Enter to exit"
    exit 1
}

# Build distinct game list + nice labels
$gamesGrouped = $allMods | Select-Object GameName -Unique | Sort-Object GameName

$games = @()
$idx   = 1
foreach ($g in $gamesGrouped) {
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
$gameChoice = Read-Host "Enter a game number to scope to that game, or press Enter for ALL games"

$selectedGame = $null
if ([string]::IsNullOrWhiteSpace($gameChoice)) {
    $modsForScope = $allMods
    $gameLabel    = "AllGames"
} else {
    if (-not ($gameChoice -as [int])) {
        Write-Host "Invalid selection. Please enter a number next time." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    $choiceInt = [int]$gameChoice
    $selected  = $games | Where-Object { $_.Index -eq $choiceInt }

    if (-not $selected) {
        Write-Host "No game with that number. Aborting." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }

    $selectedGame = $selected.GameName
    $modsForScope = $allMods | Where-Object { $_.GameName -eq $selectedGame }
    $gameLabel    = $selected.DisplayName

    if (-not $modsForScope -or $modsForScope.Count -eq 0) {
        Write-Host "No mods found for game '$($selected.DisplayName)'." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "What would you like to export?" -ForegroundColor Cyan
Write-Host "1) Installed / managed mods from Vortex (current state)"
Write-Host "2) Download archives from Vortex 'downloads' folder"
$exportMode = Read-Host "Choose 1 or 2"

if ($exportMode -eq "2") {
    # ------------------------------------------------------
    #  Download archive export
    # ------------------------------------------------------
    $archives = Get-UVMEArchiveOverview -Snapshot $snapshot -GameFilter $selectedGame

    if (-not $archives -or $archives.Count -eq 0) {
        Write-Host ""
        Write-Host "No archive files found in the Vortex download folder(s) for the chosen scope." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 0
    }

    Write-Host ""
    Write-Host "1) Export as CSV"
    Write-Host "2) Export as JSON"
    Write-Host "3) Export as Excel (XLSX)"
    Write-Host "4) Export as HTML (sortable, with full paths)"
    $formatChoice = Read-Host "Choose 1, 2, 3 or 4"

    $basePath = $PSScriptRoot
    if (-not $basePath) { $basePath = (Get-Location).Path }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileBase  = if ($selectedGame) { "$(Format-GameLabel $selectedGame)_Downloads_$timestamp" } else { "AllGames_Downloads_$timestamp" }

    $sortedArchives = $archives | Sort-Object GameName, FileName

    switch ($formatChoice) {
        "1" {
            $outFile = Join-Path $basePath "$fileBase.csv"
            $sortedArchives |
                Select-Object GameName, FileName, ModName, SizeMB, Modified, ExistsOnDisk, FullPath |
                Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
            Write-Host "`nCSV exported to:`n  $outFile" -ForegroundColor Green
        }
        "2" {
            $outFile = Join-Path $basePath "$fileBase.json"
            $sortedArchives |
                Select-Object GameName, FileName, ModName, SizeMB, Modified, ExistsOnDisk, FullPath |
                ConvertTo-Json -Depth 6 |
                Out-File -FilePath $outFile -Encoding UTF8
            Write-Host "`nJSON exported to:`n  $outFile" -ForegroundColor Green
        }
        "3" {
            $outFile = Join-Path $basePath "$fileBase.xlsx"

            $excel         = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook      = $excel.Workbooks.Add()
            $sheet         = $workbook.Worksheets.Item(1)

            $headers = @("Game","File Name","Mod Name","Size (MB)","Modified","Exists","Full Path")

            for ($c = 0; $c -lt $headers.Count; $c++) {
                $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
            }

            $row = 2
            foreach ($a in $sortedArchives) {
                $sheet.Cells.Item($row, 1).Value2 = [string]$a.GameName
                $sheet.Cells.Item($row, 2).Value2 = [string]$a.FileName
                $sheet.Cells.Item($row, 3).Value2 = [string]$a.ModName
                $sheet.Cells.Item($row, 4).Value2 = [string]$a.SizeMB
                $sheet.Cells.Item($row, 5).Value2 = if ($a.Modified) { [string]$a.Modified } else { "" }
                $sheet.Cells.Item($row, 6).Value2 = [string]$a.ExistsOnDisk
                $sheet.Cells.Item($row, 7).Value2 = [string]$a.FullPath
                $row++
            }

            New-ExcelLayout -Sheet $sheet

            $workbook.SaveAs($outFile)
            $workbook.Close($true)
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null

            Write-Host "`nExcel file exported to:`n  $outFile" -ForegroundColor Green
        }
        "4" {
            $outFile = Join-Path $basePath "$fileBase.html"
            $title   = "UVME - Download Archives - $gameLabel"

            $rowsHtml = foreach ($group in ($sortedArchives | Group-Object GameName | Sort-Object Name)) {
                $gameNice = Format-GameLabel $group.Name

                $innerRows = foreach ($a in $group.Group) {
                    $gameName = [System.Net.WebUtility]::HtmlEncode([string]$a.GameName)
                    $fileName = [System.Net.WebUtility]::HtmlEncode([string]$a.FileName)
                    $modName  = [System.Net.WebUtility]::HtmlEncode([string]$a.ModName)
                    $sizeMB   = [System.Net.WebUtility]::HtmlEncode([string]$a.SizeMB)
                    $modified = if ($a.Modified) { [System.Net.WebUtility]::HtmlEncode([string]$a.Modified) } else { "" }
                    $exists   = [System.Net.WebUtility]::HtmlEncode([string]$a.ExistsOnDisk)
                    $fullPath = [System.Net.WebUtility]::HtmlEncode([string]$a.FullPath)

                    "<tr>
                        <td>$gameName</td>
                        <td>$fileName</td>
                        <td>$modName</td>
                        <td>$sizeMB</td>
                        <td>$modified</td>
                        <td>$exists</td>
                        <td>$fullPath</td>
                    </tr>"
                }

                "<h2>$gameNice</h2>
                 <div class=""meta"">Archive files in this game: $($group.Count)</div>
                 <table class=""uvme-table"">
                   <thead>
                     <tr>
                       <th class=""sortable"" data-col=""0"" data-type=""text"">Game (sort)</th>
                       <th class=""sortable"" data-col=""1"" data-type=""text"">File Name (sort)</th>
                       <th class=""sortable"" data-col=""2"" data-type=""text"">Mod Name (sort)</th>
                       <th class=""sortable"" data-col=""3"" data-type=""number"">Size (MB) (sort)</th>
                       <th class=""sortable"" data-col=""4"" data-type=""text"">Modified (sort)</th>
                       <th class=""sortable"" data-col=""5"" data-type=""text"">Exists On Disk (sort)</th>
                       <th class=""sortable"" data-col=""6"" data-type=""text"">Full Path (sort)</th>
                     </tr>
                   </thead>
                   <tbody>
                     $($innerRows -join "`r`n")
                   </tbody>
                 </table>"
            }

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
            font-size: 24px;
            margin-bottom: 4px;
        }
        h2 {
            font-size: 18px;
            margin-top: 24px;
            margin-bottom: 4px;
        }
        .meta {
            font-size: 11px;
            color: #a1a1aa;
            margin-bottom: 6px;
        }
        table.uvme-table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 8px;
            font-size: 13px;
        }
        table.uvme-table th, table.uvme-table td {
            border: 1px solid #3f3f46;
            padding: 6px 8px;
            vertical-align: top;
        }
        table.uvme-table th {
            background-color: #27272a;
            position: sticky;
            top: 0;
            z-index: 1;
        }
        table.uvme-table tr:nth-child(even) {
            background-color: #18181b;
        }
        table.uvme-table tr:nth-child(odd) {
            background-color: #09090b;
        }
        a {
            color: #22c55e;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        .sortable {
            cursor: pointer;
        }
    </style>
</head>
<body>

    <h1>Universal Vortex Mod Exporter</h1>
    <h2>$gameLabel - Download Archives</h2>

    <div class="meta">
        Generated: $(Get-Date)<br />
        Total archive entries: $($sortedArchives.Count)
    </div>

    $($rowsHtml -join "`r`n")

    <script>
      (function () {
        function compare(a, b, type) {
          if (type === 'number') {
            var na = parseFloat(a) || 0;
            var nb = parseFloat(b) || 0;
            return na - nb;
          }
          var ta = a.toString().toLowerCase();
          var tb = b.toString().toLowerCase();
          if (ta < tb) return -1;
          if (ta > tb) return 1;
          return 0;
        }

        document.querySelectorAll('th.sortable').forEach(function (th) {
          th.addEventListener('click', function () {
            var table = th.closest('table');
            var tbody = table.querySelector('tbody');
            var col   = parseInt(th.getAttribute('data-col'), 10);
            var type  = th.getAttribute('data-type') || 'text';
            var rows  = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
            var asc   = !th.classList.contains('sorted-asc');

            table.querySelectorAll('th.sortable').forEach(function (h) {
              h.classList.remove('sorted-asc', 'sorted-desc');
            });
            th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');

            rows.sort(function (r1, r2) {
              var c1 = r1.children[col].textContent.trim();
              var c2 = r2.children[col].textContent.trim();
              var cmp = compare(c1, c2, type);
              return asc ? cmp : -cmp;
            });

            rows.forEach(function (r) { tbody.appendChild(r); });
          });
        });
      })();
    </script>

</body>
</html>
"@

            $html | Set-Content -Path $outFile -Encoding UTF8
            Write-Host "`nHTML exported to (open this in your browser):`n  $outFile" -ForegroundColor Green
        }
        default {
            Write-Host "Invalid choice. Aborting." -ForegroundColor Red
        }
    }

    Write-Host "`nDone." -ForegroundColor Cyan
    Read-Host "Press Enter to exit"
    exit 0
}

# ----------------------------------------------------------
#  Installed / managed mods export
# ----------------------------------------------------------

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
        Write-Host "Invalid choice. Aborting." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

if (-not $modsToExport -or $modsToExport.Count -eq 0) {
    Write-Host "No mods matching selection (maybe no enabled mods?). Aborting." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "1) Export as CSV"
Write-Host "2) Export as JSON"
Write-Host "3) Export as Excel (XLSX with clickable links)"
Write-Host "4) Export as HTML (grouped by game, sortable, with clickable links)"
$formatChoice = Read-Host "Choose 1, 2, 3 or 4"

$basePath = $PSScriptRoot
if (-not $basePath) { $basePath = (Get-Location).Path }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileBase  = "${gameLabel}_${scopeLabel}_${timestamp}"

$sortedMods  = $modsToExport | Sort-Object GameName, {[int]$_.LoadOrder}, ModName
$selectProps = "GameName","ModName","ModVersion","Enabled","LoadOrder","Source","Homepage"

switch ($formatChoice) {
    "1" {
        $outFile = Join-Path $basePath "$fileBase.csv"
        $sortedMods |
            Select-Object $selectProps |
            Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
        Write-Host "`nCSV exported to:`n  $outFile" -ForegroundColor Green
    }
    "2" {
        $outFile = Join-Path $basePath "$fileBase.json"
        $sortedMods |
            Select-Object $selectProps |
            ConvertTo-Json -Depth 5 |
            Out-File -FilePath $outFile -Encoding UTF8
        Write-Host "`nJSON exported to:`n  $outFile" -ForegroundColor Green
    }
    "3" {
        $outFile = Join-Path $basePath "$fileBase.xlsx"

        $excel         = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook      = $excel.Workbooks.Add()
        $sheet         = $workbook.Worksheets.Item(1)

        $headers = @("Game","Mod Name","Version","Enabled","Load Order","Source","Homepage")

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

            $sourceText = if ($m.Source -is [string] -and $m.Source.Trim()) { $m.Source } else { "no vortex record" }
            $sheet.Cells.Item($row, 6).Value2 = [string]$sourceText

            $url = $m.Homepage
            if (Test-UVMEIsUrl $url) {
                $cell = $sheet.Cells.Item($row, 7)
                $cell.Value2 = "Download"
                $sheet.Hyperlinks.Add($cell, $url, "", "", "Download") | Out-Null
            } else {
                $sheet.Cells.Item($row, 7).Value2 = [string]("No download link")
            }

            $row++
        }

        New-ExcelLayout -Sheet $sheet

        $workbook.SaveAs($outFile)
        $workbook.Close($true)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null

        Write-Host "`nExcel file exported to:`n  $outFile" -ForegroundColor Green
    }
    "4" {
        $outFile = Join-Path $basePath "$fileBase.html"

        $title = "Universal Vortex Mod Exporter - $gameLabel ($scopeLabel)"
        $bannerImg = "images/uvme-banner.png"
        $avatarImg = "images/vaultboy.png"

        $gamesForHtml = $sortedMods | Group-Object GameName | Sort-Object Name

        $sections = foreach ($group in $gamesForHtml) {
            $gameId   = $group.Name
            $gameNice = Format-GameLabel $gameId
            $rowsHtml = foreach ($m in $group.Group) {
                $gameName   = [System.Net.WebUtility]::HtmlEncode([string]$m.GameName)
                $modName    = [System.Net.WebUtility]::HtmlEncode([string]$m.ModName)
                $modVersion = [System.Net.WebUtility]::HtmlEncode([string]$m.ModVersion)
                $enabled    = [System.Net.WebUtility]::HtmlEncode([string]$m.Enabled)
                $loadOrder  = [System.Net.WebUtility]::HtmlEncode([string]$m.LoadOrder)

                $sourceText = if ($m.Source -is [string] -and $m.Source.Trim()) {
                    $m.Source
                } else {
                    "no vortex record"
                }
                $source     = [System.Net.WebUtility]::HtmlEncode([string]$sourceText)

                $url = [string]$m.Homepage
                if (Test-UVMEIsUrl $url) {
                    $urlEsc = [System.Net.WebUtility]::HtmlEncode($url)
                    $homepageCell = "<a href=""$urlEsc"" target=""_blank"" rel=""noopener"">Download</a>"
                } else {
                    $homepageCell = "No download link"
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

            "<h2>$gameNice</h2>
             <div class=""meta"">Mods in this game: $($group.Count)</div>
             <table class=""uvme-table"">
               <thead>
                 <tr>
                   <th class=""sortable"" data-col=""0"" data-type=""text"">Game (sort)</th>
                   <th class=""sortable"" data-col=""1"" data-type=""text"">Mod Name (sort)</th>
                   <th class=""sortable"" data-col=""2"" data-type=""text"">Version (sort)</th>
                   <th class=""sortable"" data-col=""3"" data-type=""text"">Enabled (sort)</th>
                   <th class=""sortable"" data-col=""4"" data-type=""number"">Load Order (sort)</th>
                   <th class=""sortable"" data-col=""5"" data-type=""text"">Source (sort)</th>
                   <th class=""sortable"" data-col=""6"" data-type=""text"">Homepage (sort)</th>
                 </tr>
               </thead>
               <tbody>
                 $($rowsHtml -join "`r`n")
               </tbody>
             </table>"
        }

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
            font-size: 24px;
            margin-bottom: 4px;
        }
        h2 {
            font-size: 18px;
            margin-top: 24px;
            margin-bottom: 4px;
        }
        .meta {
            font-size: 11px;
            color: #a1a1aa;
            margin-bottom: 6px;
        }

        .header-visual {
            display: flex;
            justify-content: center;
            align-items: flex-end;
            gap: 16px;
            margin-bottom: 12px;
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

        table.uvme-table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 8px;
            font-size: 13px;
        }
        table.uvme-table th, table.uvme-table td {
            border: 1px solid #3f3f46;
            padding: 6px 8px;
            vertical-align: top;
        }
        table.uvme-table th {
            background-color: #27272a;
            position: sticky;
            top: 0;
            z-index: 1;
        }
        table.uvme-table tr:nth-child(even) {
            background-color: #18181b;
        }
        table.uvme-table tr:nth-child(odd) {
            background-color: #09090b;
        }
        a {
            color: #22c55e;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        .sortable {
            cursor: pointer;
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
    <h2>$gameLabel - $scopeLabel</h2>

    <div class="meta">
        Generated: $(Get-Date)<br />
        Total mods: $($sortedMods.Count)
    </div>

    $($sections -join "`r`n")

    <script>
      (function () {
        function compare(a, b, type) {
          if (type === 'number') {
            var na = parseFloat(a) || 0;
            var nb = parseFloat(b) || 0;
            return na - nb;
          }
          var ta = a.toString().toLowerCase();
          var tb = b.toString().toLowerCase();
          if (ta < tb) return -1;
          if (ta > tb) return 1;
          return 0;
        }

        document.querySelectorAll('th.sortable').forEach(function (th) {
          th.addEventListener('click', function () {
            var table = th.closest('table');
            var tbody = table.querySelector('tbody');
            var col   = parseInt(th.getAttribute('data-col'), 10);
            var type  = th.getAttribute('data-type') || 'text';
            var rows  = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
            var asc   = !th.classList.contains('sorted-asc');

            table.querySelectorAll('th.sortable').forEach(function (h) {
              h.classList.remove('sorted-asc', 'sorted-desc');
            });
            th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');

            rows.sort(function (r1, r2) {
              var c1 = r1.children[col].textContent.trim();
              var c2 = r2.children[col].textContent.trim();
              var cmp = compare(c1, c2, type);
              return asc ? cmp : -cmp;
            });

            rows.forEach(function (r) { tbody.appendChild(r); });
          });
        });
      })();
    </script>

</body>
</html>
"@

        $html | Set-Content -Path $outFile -Encoding UTF8
        Write-Host "`nHTML exported to (open this in your browser):`n  $outFile" -ForegroundColor Green
    }
    default {
        Write-Host "Invalid choice. Aborting." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host "`nDone." -ForegroundColor Cyan
Read-Host "Press Enter to exit"
