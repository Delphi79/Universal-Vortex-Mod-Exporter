[CmdletBinding()]
param()

# ==========================================================
#  Global cache
# ==========================================================
$script:UVME_SnapshotCache   = $null
$script:UVME_DuplicatesFixed = $false

# Small helper to safely get a trimmed string
function Get-UVMEString {
    param($Value)

    if ($null -eq $Value) { return $null }
    return ([string]$Value).Trim()
}

function Fix-UVMEJsonDuplicates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Json
    )

    # Fast path: if JSON parses fine, don't touch it
    try {
        $null = $Json | ConvertFrom-Json -ErrorAction Stop
        return $Json
    } catch {
        if ($_.Exception.Message -notmatch 'Duplicate.*property name') {
            # It's failing for some other reason; let the caller see that
            throw
        }
    }

    Write-Host "UVME: Detected duplicate JSON property names. Attempting generic de-duplication..." -ForegroundColor Yellow

    $text  = $Json
    $len   = $text.Length

    $sb    = New-Object System.Text.StringBuilder
    $stack = New-Object System.Collections.Stack

    # Each object on the stack holds a case-insensitive set of keys seen at that object level
    $stack.Push(@{})

    $inString = $false
    $escape   = $false

    for ($i = 0; $i -lt $len; $i++) {
        $ch = $text[$i]

        if ($inString) {
            # Inside string: just copy and track escapes
            $sb.Append($ch) | Out-Null

            if ($escape) {
                $escape = $false
            } elseif ($ch -eq '\') {
                $escape = $true
            } elseif ($ch -eq '"') {
                $inString = $false
            }
            continue
        }

        switch ($ch) {
            '{' {
                $sb.Append($ch) | Out-Null
                # New object scope, new key set
                $stack.Push(@{})
                continue
            }
            '}' {
                $sb.Append($ch) | Out-Null
                if ($stack.Count -gt 1) {
                    $null = $stack.Pop()
                }
                continue
            }
            '"' {
                # Potential start of a key or a string value

                # Look backwards for the previous non-whitespace char
                $j = $i - 1
                while ($j -ge 0 -and [char]::IsWhiteSpace($text[$j])) { $j-- }
                $prev = if ($j -ge 0) { $text[$j] } else { [char]0 }

                # Find the closing quote for this string, respecting escapes
                $k = $i + 1
                $esc2 = $false
                while ($k -lt $len) {
                    $c2 = $text[$k]
                    if ($esc2) {
                        $esc2 = $false
                    } elseif ($c2 -eq '\') {
                        $esc2 = $true
                    } elseif ($c2 -eq '"') {
                        break
                    }
                    $k++
                }

                if ($k -ge $len) {
                    # Malformed JSON – just fall back to normal string handling
                    $inString = $true
                    $sb.Append($ch) | Out-Null
                    continue
                }

                # Look forward to see if this string is followed by a colon => it's a property name
                $l = $k + 1
                while ($l -lt $len -and [char]::IsWhiteSpace($text[$l])) { $l++ }
                $next = if ($l -lt $len) { $text[$l] } else { [char]0 }

                $isKey = ($prev -eq '{' -or $prev -eq ',') -and $next -eq ':'

                if ($isKey -and $stack.Count -gt 0) {
                    # This is a property name at the current object depth
                    $keyText = $text.Substring($i + 1, $k - $i - 1)

                    $ctx = $stack.Peek()
                    $keyNorm = $keyText.ToLowerInvariant()

                    if ($ctx.ContainsKey($keyNorm)) {
                        # Duplicate key in this object: rename this occurrence
                        $baseName = $keyText
                        $suffixIndex = 1
                        $newName = $null

                        while ($true) {
                            $candidate = if ($suffixIndex -eq 1) {
                                "$baseName (Duplicate)"
                            } else {
                                "$baseName (Duplicate $suffixIndex)"
                            }

                            $candNorm = $candidate.ToLowerInvariant()
                            if (-not $ctx.ContainsKey($candNorm)) {
                                $newName = $candidate
                                $ctx[$candNorm] = $true
                                break
                            }
                            $suffixIndex++
                        }

                        Write-Host "UVME: Renaming duplicate JSON key '$keyText' => '$newName'." -ForegroundColor Yellow

                        # Write the renamed key (with quotes)
                        $sb.Append('"').Append($newName).Append('"') | Out-Null
                    } else {
                        # First time we've seen this key at this depth
                        $ctx[$keyNorm] = $true

                        # Copy original key string including quotes
                        $sb.Append($text, $i, $k - $i + 1) | Out-Null
                    }

                    # Skip the characters we just handled inside the quotes
                    $i = $k
                    continue
                } else {
                    # Normal string value, not a key
                    $inString = $true
                    $sb.Append($ch) | Out-Null
                    continue
                }
            }
            default {
                $sb.Append($ch) | Out-Null
                continue
            }
        }
    }

    $fixed = $sb.ToString()

    # Try again after fixing duplicates
    try {
        $null = $fixed | ConvertFrom-Json -ErrorAction Stop
        Write-Host "UVME: JSON de-duplication complete." -ForegroundColor Yellow
        return $fixed
    } catch {
        Write-Host "UVME: JSON still failed to parse after de-duplication: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# ==========================================================
#  Helpers
# ==========================================================

function Format-GameLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $t = $Name -replace '_', ' '

    if ($t -match '^(.*?)(\d+)$') {
        $t = "$($matches[1]) $($matches[2])"
    }

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $label   = $culture.TextInfo.ToTitleCase($t.ToLower())

    return $label.Trim()
}

function Get-VortexSnapshot {
    if ($script:UVME_SnapshotCache) {
        return $script:UVME_SnapshotCache
    }

    $pattern = Join-Path $env:APPDATA "Vortex\temp\state_backups_full\*.json"

    $latest = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

    if (-not $latest) {
        throw "No Vortex backup JSON found at $pattern. Start/restart Vortex so it writes a full state backup, then run UVME again."
    }

    $rawJson = Get-Content -Path $latest.FullName -Raw

    # Generic duplicate-key handling (for any game / section)
    if (-not $script:UVME_DuplicatesFixed) {
        $rawJson = Fix-UVMEJsonDuplicates -Json $rawJson
        $script:UVME_DuplicatesFixed = $true
    }

    $snapshot = $rawJson | ConvertFrom-Json
    $script:UVME_SnapshotCache = $snapshot
    return $snapshot
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
        $gameName  = $gameProp.Name
        $modBlock  = $gameProp.Value
        if (-not $modBlock) { continue }

        $profileId = $null
        if ($lastActiveMap -and $lastActiveMap.PSObject.Properties.Name -contains $gameName) {
            $profileId = $lastActiveMap.$gameName
        }

        $profileState = $null
        if ($profileId -and $profilesRoot -and $profilesRoot.PSObject.Properties.Name -contains $profileId) {
            $profileState = $profilesRoot.$profileId.modState
        }

        $deployIndex = 0
        foreach ($modProp in $modBlock.PSObject.Properties) {
            $modKey = $modProp.Name
            $m      = $modProp.Value
            if (-not $m) {
                $deployIndex++
                continue
            }

            $attrs = $m.attributes

            # Enabled flag from profile
            $enabled = $false
            if ($profileState -and $profileState.PSObject.Properties.Name -contains $modKey) {
                $enabled = [bool]$profileState.$modKey.enabled
            }

            # Raw attribute values from Vortex
            $rawLogical = $null
            $rawModName = $null
            $rawName    = $null

            if ($attrs) {
                if ($attrs.logicalFileName) {
                    $tmp = Get-UVMEString $attrs.logicalFileName
                    if ($tmp.Length -gt 0) { $rawLogical = $tmp }
                }
                if ($attrs.modName) {
                    $tmp = Get-UVMEString $attrs.modName
                    if ($tmp.Length -gt 0) { $rawModName = $tmp }
                }
                if ($attrs.name) {
                    $tmp = Get-UVMEString $attrs.name
                    if ($tmp.Length -gt 0) { $rawName = $tmp }
                }
            }

            # --- Vortex-like display name selection ---

            if ($rawLogical) {
                # 1) Prefer the same field Vortex shows in the Mods list
                $modName = Get-UVMEString $rawLogical
            }
            elseif ($rawModName) {
                # 2) Fall back to mod/page name
                $modName = Get-UVMEString $rawModName
            }
            elseif ($rawName) {
                # 3) Last resort: archive-ish name, cleaned up
                $clean = Get-UVMEString $rawName

                # Strip patterns like "-5124-3-09-1739477203" at the end
                $clean = $clean -replace '-\d+-\d+(?:-\d+)*-\d{9,}$',''

                # Strip common archive extensions
                $clean = $clean -replace '\.(zip|rar|7z|7zip)$',''

                $modName = Get-UVMEString $clean
            }
            else {
                # Fallback when literally nothing is usable
                if ($m.type) {
                    $typeStr = Get-UVMEString $m.type
                    if ($typeStr.Length -gt 0) {
                        $modName = "[Tool entry - $typeStr]"
                    } else {
                        $modName = "[Unnamed entry - Vortex has no mod name]"
                    }
                } else {
                    $modName = "[Unnamed entry - Vortex has no mod name]"
                }
            }

            # BaseModName used for grouping: prefer mod/page name, then final display name
            $baseName = Get-UVMEString $rawModName
            if (-not $baseName -or $baseName.Length -eq 0) {
                $baseName = Get-UVMEString $modName
            }

            # Version: prefer per-file version, then global modVersion
            $fileVersion   = $null
            $globalVersion = $null
            $modVersion    = $null

            if ($attrs) {
                if ($attrs.version) {
                    $tmp = Get-UVMEString $attrs.version
                    if ($tmp.Length -gt 0) { $fileVersion = $tmp }
                }
                if ($attrs.modVersion) {
                    $tmp = Get-UVMEString $attrs.modVersion
                    if ($tmp.Length -gt 0) { $globalVersion = $tmp }
                }
            }

            if ($fileVersion) {
                $modVersion = $fileVersion
            } elseif ($globalVersion) {
                $modVersion = $globalVersion
            } else {
                $modVersion = "[no version in Vortex]"
            }

            # Source + homepage + Nexus ID
            $source       = $null
            $homepage     = $null
            $modNumericId = $null
            $downloadGame = $null

            if ($attrs) {
                if ($attrs.PSObject.Properties.Name -contains 'source') {
                    $tmp = Get-UVMEString $attrs.source
                    if ($tmp.Length -gt 0) { $source = $tmp }
                }
                if ($attrs.PSObject.Properties.Name -contains 'homepage') {
                    $tmp = Get-UVMEString $attrs.homepage
                    if ($tmp.Length -gt 0) { $homepage = $tmp }
                }
                if ($attrs.PSObject.Properties.Name -contains 'modId') {
                    $modNumericId = [string]$attrs.modId
                }
                if ($attrs.PSObject.Properties.Name -contains 'downloadGame') {
                    $downloadGame = Get-UVMEString $attrs.downloadGame
                }
            }

            $rows += [PSCustomObject]@{
                GameName       = $gameName
                ModId          = $m.id
                ModKey         = $modKey
                ModName        = $modName
                BaseModName    = $baseName
                ModVersion     = $modVersion
                Enabled        = $enabled
                LoadOrder      = [string]$deployIndex
                Source         = $source
                Homepage       = $homepage
                FileVersion    = $fileVersion
                GlobalVersion  = $globalVersion
                RawLogicalName = $rawLogical
                RawModName     = $rawModName
                RawName        = $rawName
                ModNumericId   = $modNumericId
                DownloadGame   = $downloadGame
            }

            $deployIndex++
        }
    }

    return $rows
}

function Normalize-UVMEPartMods {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Mods
    )

    $modsArray = @()
    foreach ($m in $Mods) { $modsArray += $m }

    $result = @()

    # Group by Game + Homepage (Nexus page) for per-mod grouping
    $groups = $modsArray | Group-Object GameName, Homepage

    foreach ($g in $groups) {
        $items = $g.Group
        if ($items.Count -eq 0) { continue }

        # Detect "Part N" entries in this group
        $partInfos = @()

        foreach ($row in $items) {
            $partNum = $null
            $candidates = @(
                $row.RawLogicalName,
                $row.RawModName,
                $row.RawName,
                $row.ModName,
                $row.ModKey
            ) | Where-Object { $_ -and (Get-UVMEString $_).Length -gt 0 }

            foreach ($txt in $candidates) {
                $txtStr = Get-UVMEString $txt
                if ($txtStr -match '(?i)\(Part\s+0*([0-9]+)\)') {
                    $partNum = [int]$matches[1]
                    break
                } elseif ($txtStr -match '(?i)\bpart\s+0*([0-9]+)\b') {
                    $partNum = [int]$matches[1]
                    break
                }
            }

            if ($partNum -ne $null) {
                $partInfos += [PSCustomObject]@{
                    Row     = $row
                    PartNum = $partNum
                }
            }
        }

        if ($partInfos.Count -lt 2) {
            # Not clearly multi-part -> keep all rows as-is
            $result += $items
            continue
        }

        # We have a multi-part mod: derive a base name
        $baseCandidates = @()

        foreach ($row in $items) {
            $nameFields = @(
                $row.RawModName,
                $row.RawLogicalName,
                $row.RawName,
                $row.BaseModName,
                $row.ModName
            ) | Where-Object { $_ -and (Get-UVMEString $_).Length -gt 0 }

            foreach ($txt in $nameFields) {
                $t = Get-UVMEString $txt
                if ($t -match '(?i)^(.*?)(?:\s*[-–]\s*)?part\s+0*[0-9]+.*$') {
                    $base = Get-UVMEString $matches[1]
                    if ($base.Length -gt 2) {
                        $baseCandidates += $base
                    }
                }
            }
        }

        if ($baseCandidates.Count -eq 0) {
            $baseCandidates = $items |
                ForEach-Object { $_.BaseModName, $_.RawModName, $_.RawName, $_.ModName } |
                Where-Object { $_ -and (Get-UVMEString $_).Length -gt 0 } |
                Sort-Object { (Get-UVMEString $_).Length } -Descending -Unique
        }

        if ($baseCandidates.Count -eq 0) {
            # Can't confidently determine base name; keep rows as-is
            $result += $items
            continue
        }

        $baseName = Get-UVMEString $baseCandidates[0]
        if ($baseName -match '(?i)^(.*?)(?:\s*[-–]\s*)?part\s+0*[0-9]+.*$') {
            $baseName = Get-UVMEString $matches[1]
        }
        if (-not $baseName -or $baseName.Length -le 2) {
            $result += $items
            continue
        }

        # Determine unified version for the mod
        $globalVers = $partInfos |
            ForEach-Object { Get-UVMEString $_.Row.GlobalVersion } |
            Where-Object { $_ -and $_.Length -gt 0 } |
            Select-Object -Unique

        $commonVersion = $null
        if ($globalVers.Count -ge 1) {
            $commonVersion = $globalVers[0]
        } else {
            $freq = $partInfos |
                Group-Object { Get-UVMEString $_.Row.ModVersion } |
                Sort-Object Count -Descending

            if ($freq.Count -ge 1) {
                $commonVersion = Get-UVMEString $freq[0].Name
            }
        }
        if (-not $commonVersion -or -not (Get-UVMEString $commonVersion)) {
            $commonVersion = $items[0].ModVersion
        }

        # Aggregate Enabled + LoadOrder
        $anyEnabled  = ($items | Where-Object { $_.Enabled }) -ne $null

        # Only consider rows with a numeric load order, then pick the smallest
        $itemsWithNumericLoad = $items | Where-Object {
            $_.LoadOrder -ne $null -and
            $_.LoadOrder.ToString().Trim() -match '^\d+$'
        }

        if ($itemsWithNumericLoad.Count -gt 0) {
            $representative = $itemsWithNumericLoad |
                Sort-Object { [int]$_.LoadOrder } |
                Select-Object -First 1
            $minLoad = [int]$representative.LoadOrder
        } else {
            # Fallback if all load orders are missing/invalid
            $representative = $items | Select-Object -First 1
            $minLoad = 0
        }

        # Use any homepage present within this multi-part group
        $groupHomepages = $items |
            ForEach-Object { Get-UVMEString $_.Homepage } |
            Where-Object { $_ -and $_.Length -gt 0 } |
            Select-Object -Unique

        $aggHomepage = if ($groupHomepages.Count -gt 0) { $groupHomepages[0] } else { $representative.Homepage }

        $agg = [PSCustomObject]@{
            GameName       = $representative.GameName
            ModId          = $representative.ModId
            ModKey         = $representative.ModKey
            ModName        = $baseName
            BaseModName    = $baseName
            ModVersion     = $commonVersion
            Enabled        = [bool]$anyEnabled
            LoadOrder      = [string]$minLoad
            Source         = $representative.Source
            Homepage       = $aggHomepage
            FileVersion    = $representative.FileVersion
            GlobalVersion  = $representative.GlobalVersion
            RawLogicalName = $representative.RawLogicalName
            RawModName     = $representative.RawModName
            RawName        = $representative.RawName
            ModNumericId   = $representative.ModNumericId
            DownloadGame   = $representative.DownloadGame
        }

        # Only one row for this multi-part mod
        $result += $agg
    }

    return $result
}

function Normalize-UVMEGenericLabels {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Mods
    )

    $modsArray = @()
    foreach ($m in $Mods) { $modsArray += $m }

    $groups = $modsArray | Group-Object GameName, RawLogicalName

    foreach ($g in $groups) {
        $items = $g.Group
        if ($items.Count -lt 2) { continue }

        $logical = Get-UVMEString $items[0].RawLogicalName
        if (-not $logical -or $logical.Length -eq 0) { continue }

        # Distinct mod names or homepages?
        $distinctModNames = $items |
            ForEach-Object { Get-UVMEString $_.RawModName } |
            Where-Object { $_ -and $_.Length -gt 0 } |
            Select-Object -Unique

        $distinctHomes = $items |
            ForEach-Object { Get-UVMEString $_.Homepage } |
            Where-Object { $_ -and $_.Length -gt 0 } |
            Select-Object -Unique

        if ($distinctModNames.Count -lt 2 -and $distinctHomes.Count -lt 2) {
            continue
        }

        # Promote mod/page name (or archive name) so rows are distinguishable
        foreach ($row in $items) {
            $candidate = Get-UVMEString $row.RawModName
            if (-not $candidate -or $candidate.Length -eq 0) {
                $candidate = Get-UVMEString $row.RawName
            }

            if ($candidate -and $candidate.Length -gt 0) {
                $row.ModName     = $candidate
                $row.BaseModName = $candidate
            }
        }
    }

    return $modsArray
}

function Normalize-UVMEHomepages {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Mods
    )

    # Materialize into an array so we can do multi-pass logic
    $modsArray = @()
    foreach ($m in $Mods) { $modsArray += $m }

    #
    # STEP 1: Propagate homepage within (GameName, ModNumericId) groups
    #         e.g. all SBS files share modId = 84015
    #
    $groups = $modsArray |
        Where-Object {
            $_.ModNumericId -and (Get-UVMEString $_.ModNumericId).Length -gt 0
        } |
        Group-Object GameName, ModNumericId

    foreach ($g in $groups) {
        $items = $g.Group
        if ($items.Count -lt 2) { continue }

        # Any homepage in this modId group?
        $homepages = $items |
            ForEach-Object { Get-UVMEString $_.Homepage } |
            Where-Object { $_ -and $_.Length -gt 0 } |
            Select-Object -Unique

        if ($homepages.Count -lt 1) { continue }

        $chosen = $homepages[0]

        foreach ($row in $items) {
            $cur = Get-UVMEString $row.Homepage
            if (-not $cur -or $cur.Length -eq 0) {
                $row.Homepage = $chosen
            }
        }
    }

    #
    # STEP 2: Learn Nexus "game slug" per GameName from existing homepages
    #         Example:
    #         https://www.nexusmods.com/fallout4/mods/84015/
    #         https://www.nexusmods.com/skyrimspecialedition/mods/32444/
    #         -> slugs: fallout4, skyrimspecialedition
    #
    $slugByGame = @{}

    foreach ($row in $modsArray) {
        $src = Get-UVMEString $row.Source
        $hp  = Get-UVMEString $row.Homepage

        if (-not $src -or $src.ToLower() -ne 'nexus') { continue }
        if (-not (Test-UVMEIsUrl $hp)) { continue }

        $m = [regex]::Match(
            $hp,
            '^https?://(?:www\.)?nexusmods\.com/([^/]+)/mods/\d+/?'
        )
        if ($m.Success) {
            $gameName = $row.GameName
            if ($gameName -and -not $slugByGame.ContainsKey($gameName)) {
                $slugByGame[$gameName] = $m.Groups[1].Value
            }
        }
    }

    #
    # STEP 3: For any remaining Nexus mods with a numeric modId but no homepage,
    #         synthesize the URL from (slug, modId).
    #
    foreach ($row in $modsArray) {
        $src = Get-UVMEString $row.Source
        if (-not $src -or $src.ToLower() -ne 'nexus') { continue }

        $hp = Get-UVMEString $row.Homepage
        if (Test-UVMEIsUrl $hp) { continue }  # already has a real URL

        $modIdStr = Get-UVMEString $row.ModNumericId
        if (-not $modIdStr -or -not ($modIdStr -match '^\d+$')) { continue }

        $gameName = $row.GameName
        if (-not $gameName) { continue }
        if (-not $slugByGame.ContainsKey($gameName)) { continue }

        $slug = $slugByGame[$gameName]
        $row.Homepage = "https://www.nexusmods.com/$slug/mods/$modIdStr/"
    }

    return $modsArray
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
                $modName = Get-UVMEString $entry.modInfo.name
            }

            $sizeBytes = [double]($entry.size | ForEach-Object { $_ })
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
                GameName     = $gameName
                FileName     = $fileName
                ModName      = $modName
                SizeMB       = $sizeMB
                Modified     = $modified
                ExistsOnDisk = $exists
                FullPath     = $fullPath
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
        $ps.Orientation = 2
        $ps.Zoom        = 100
    } catch { }

    try {
        $Sheet.Columns.Item(1).ColumnWidth = 12
        $Sheet.Columns.Item(2).ColumnWidth = 45
        $Sheet.Columns.Item(3).ColumnWidth = 18
        $Sheet.Columns.Item(4).ColumnWidth = 10
        $Sheet.Columns.Item(5).ColumnWidth = 12
        $Sheet.Columns.Item(6).ColumnWidth = 12
        $Sheet.Columns.Item(7).ColumnWidth = 70
        $Sheet.Columns.Item(2).WrapText    = $true
    } catch { }
}

function Test-UVMEIsUrl {
    param(
        [AllowNull()]
        $Value
    )

    if (-not $Value) { return $false }
    $s = Get-UVMEString $Value
    if (-not $s) { return $false }

    return ($s.StartsWith("http://") -or $s.StartsWith("https://") -or $s.StartsWith("ftp://"))
}

# ---------- New helpers for robust user input ----------

function Read-UVMEChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string[]]$ValidValues
    )

    while ($true) {
        $choice = Read-Host $Prompt

        if ($ValidValues -contains $choice) {
            return $choice
        }

        Write-Host "Invalid selection. Please choose one of: $($ValidValues -join ', ')." -ForegroundColor Red
    }
}

function Read-UVMEGameSelection {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Games
    )

    while ($true) {
        $raw = Read-Host "Enter a game number to scope to that game, or press Enter for ALL games"

        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $null   # All games
        }

        if ($raw -as [int]) {
            $idx      = [int]$raw
            $selected = $Games | Where-Object { $_.Index -eq $idx }
            if ($selected) {
                return $selected
            }
        }

        Write-Host "That isn't a valid game selection. Please enter a number from 1 to $($Games.Count), or press Enter for ALL games." -ForegroundColor Red
    }
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
    exit 1
}

try {
    $allMods = Get-UVMEGameOverview   -Snapshot $snapshot
    $allMods = Normalize-UVMEPartMods -Mods $allMods
    $allMods = Normalize-UVMEGenericLabels -Mods $allMods
    $allMods = Normalize-UVMEHomepages     -Mods $allMods
} catch {
    Write-Error $_
    exit 1
}

if (-not $allMods -or $allMods.Count -eq 0) {
    Write-Error "No mods found in Vortex snapshots."
    exit 1
}

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
$selection = Read-UVMEGameSelection -Games $games

$selectedGame = $null
if ($null -eq $selection) {
    # ALL games
    $modsForScope = $allMods
    $gameLabel    = "AllGames"
} else {
    $selectedGame = $selection.GameName
    $modsForScope = $allMods | Where-Object { $_.GameName -eq $selectedGame }
    $gameLabel    = $selection.DisplayName

    if (-not $modsForScope -or $modsForScope.Count -eq 0) {
        Write-Host "No mods found for game '$($selection.DisplayName)'." -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "What would you like to export?" -ForegroundColor Cyan
Write-Host "1) Installed / managed mods from Vortex (current state)"
Write-Host "2) Download archives from Vortex 'downloads' folder"
$exportMode = Read-UVMEChoice -Prompt "Choose 1 or 2" -ValidValues @("1","2")

if ($exportMode -eq "2") {
    # ------------------------------------------------------
    #  Download archive export
    # ------------------------------------------------------
    $archives = Get-UVMEArchiveOverview -Snapshot $snapshot -GameFilter $selectedGame

    if (-not $archives -or $archives.Count -eq 0) {
        Write-Host ""
        Write-Host "No archive files found in the Vortex download folder(s) for the chosen scope." -ForegroundColor Yellow
        exit 0
    }

    Write-Host ""
    Write-Host "1) Export as JSON"
    Write-Host "2) Export as Excel (XLSX)"
    Write-Host "3) Export as HTML (sortable, with full paths)"
    $formatChoice = Read-UVMEChoice -Prompt "Choose 1, 2 or 3" -ValidValues @("1","2","3")

    $basePath = $PSScriptRoot
    if (-not $basePath) { $basePath = (Get-Location).Path }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileBase  = if ($selectedGame) { "$(Format-GameLabel $selectedGame)_Downloads_$timestamp" } else { "AllGames_Downloads_$timestamp" }

    $sortedArchives = $archives | Sort-Object GameName, FileName

    switch ($formatChoice) {
        "1" {
            $outFile = Join-Path $basePath "$fileBase.json"
            $sortedArchives |
                Select-Object GameName, FileName, ModName, SizeMB, Modified, ExistsOnDisk, FullPath |
                ConvertTo-Json -Depth 6 |
                Out-File -FilePath $outFile -Encoding UTF8
            Write-Host "`nJSON exported to:`n  $outFile" -ForegroundColor Green
        }
        "2" {
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

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null

            Write-Host "`nExcel file exported to:`n  $outFile" -ForegroundColor Green
        }
        "3" {
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
        h1 { font-size: 24px; margin-bottom: 4px; }
        h2 { font-size: 18px; margin-top: 24px; margin-bottom: 4px; }
        .meta { font-size: 11px; color: #a1a1aa; margin-bottom: 6px; }
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
        table.uvme-table tr:nth-child(even) { background-color: #18181b; }
        table.uvme-table tr:nth-child(odd)  { background-color: #09090b; }
        a { color: #22c55e; text-decoration: none; }
        a:hover { text-decoration: underline; }
        .sortable { cursor: pointer; }
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
    }

    Write-Host "`nDone." -ForegroundColor Cyan
    exit 0
}

# ----------------------------------------------------------
#  Installed / managed mods export
# ----------------------------------------------------------

Write-Host ""
Write-Host "1) Export ALL mods"
Write-Host "2) Export ONLY ENABLED mods"
$scopeChoice = Read-UVMEChoice -Prompt "Choose 1 or 2" -ValidValues @("1","2")

switch ($scopeChoice) {
    "1" {
        $modsToExport = $modsForScope
        $scopeLabel   = "AllMods"
    }
    "2" {
        $modsToExport = $modsForScope | Where-Object { $_.Enabled -eq $true }
        $scopeLabel   = "EnabledMods"
    }
}

if (-not $modsToExport -or $modsToExport.Count -eq 0) {
    Write-Host "No mods matching selection (maybe no enabled mods?). Aborting." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "1) Export as JSON"
Write-Host "2) Export as Excel (XLSX with clickable links)"
Write-Host "3) Export as HTML (grouped by game, sortable, with clickable links)"
$formatChoice = Read-UVMEChoice -Prompt "Choose 1, 2 or 3" -ValidValues @("1","2","3")

$basePath = $PSScriptRoot
if (-not $basePath) { $basePath = (Get-Location).Path }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileBase  = "${gameLabel}_${scopeLabel}_${timestamp}"

$sortedMods  = $modsToExport | Sort-Object GameName, {[int]$_.LoadOrder}, ModName
$selectProps = "GameName","ModName","ModVersion","Enabled","LoadOrder","Source","Homepage"

switch ($formatChoice) {
    "1" {
        $outFile = Join-Path $basePath "$fileBase.json"
        $sortedMods |
            Select-Object $selectProps |
            ConvertTo-Json -Depth 5 |
            Out-File -FilePath $outFile -Encoding UTF8
        Write-Host "`nJSON exported to:`n  $outFile" -ForegroundColor Green
    }
    "2" {
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
                $sheet.Cells.Item($row, 7).Value2 = [string]("No link in Vortex snapshot")
            }

            $row++
        }

        New-ExcelLayout -Sheet $sheet

        $workbook.SaveAs($outFile)
        $workbook.Close($true)
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null

        Write-Host "`nExcel file exported to:`n  $outFile" -ForegroundColor Green
    }
    "3" {
        $outFile = Join-Path $basePath "$fileBase.html"

        $title     = "Universal Vortex Mod Exporter - $gameLabel ($scopeLabel)"
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

                $sourceText = if ($m.Source -is [string] -and $m.Source.Trim()) { $m.Source } else { "no vortex record" }
                $source     = [System.Net.WebUtility]::HtmlEncode([string]$sourceText)

                $url = [string]$m.Homepage
                if (Test-UVMEIsUrl $url) {
                    $urlEsc       = [System.Net.WebUtility]::HtmlEncode($url)
                    $homepageCell = "<a href=""$urlEsc"" target=""_blank"" rel=""noopener"">Download</a>"
                } else {
                    $homepageCell = "No link in Vortex snapshot"
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
        h1 { font-size: 24px; margin-bottom: 4px; }
        h2 { font-size: 18px; margin-top: 24px; margin-bottom: 4px; }
        .meta { font-size: 11px; color: #a1a1aa; margin-bottom: 6px; }
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
        table.uvme-table tr:nth-child(even) { background-color: #18181b; }
        table.uvme-table tr:nth-child(odd)  { background-color: #09090b; }
        a { color: #22c55e; text-decoration: none; }
        a:hover { text-decoration: underline; }
        .sortable { cursor: pointer; }
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
}

Write-Host "`nDone." -ForegroundColor Cyan
exit 0
