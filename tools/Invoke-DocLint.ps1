<#
.SYNOPSIS
    Structural validation of every Markdown document in the repository.

.DESCRIPTION
    These documents were edited programmatically many times - rows inserted into
    tables, sections replaced, counters rewritten. A Markdown table that loses its
    delimiter row, or a code fence left open, renders as garbage on GitHub and in
    the wiki, and no functional test would ever notice.

    Checks performed on every .md file:

      * code fences are balanced (an unclosed ``` swallows the rest of the file)
      * every table has a delimiter row whose column count matches the header
      * every table body row has the same column count as its header
      * no stray pipe-only or delimiter-only orphan rows
      * headings are properly spaced (a heading with no blank line before it does
        not render as a heading in some parsers)
      * no duplicate heading text at the same level within one file, which would
        produce colliding anchor links
      * no unresolved link placeholders or empty link targets

.PARAMETER Quiet
    Suppress per-file progress; report failures and the summary only.

.EXAMPLE
    ./tools/Invoke-DocLint.ps1
#>

[CmdletBinding()]
param(
    [AllowNull()][string]$Root,
    [switch]$Quiet
)

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }

$fail = 0
$pass = 0
$filesChecked = 0

function Report {
    param([string]$File, [string]$Issue, [int]$Line = 0)
    $script:fail++
    $loc = if ($Line -gt 0) { ":$Line" } else { '' }
    Write-Host ("  FAIL  {0}{1}  {2}" -f $File, $loc, $Issue) -ForegroundColor Red
}

$docs = @(Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.md' -File |
    Where-Object { $_.FullName -notmatch '[\\/]dist[\\/]' -and $_.FullName -notmatch '[\\/]\.git[\\/]' } |
    Sort-Object FullName)

foreach ($doc in $docs) {
    $filesChecked++
    $rel = $doc.FullName.Substring($Root.Length).TrimStart([IO.Path]::DirectorySeparatorChar, '/')
    $lines = @(Get-Content -LiteralPath $doc.FullName)
    $fileFails = $fail

    # ---- 1. balanced code fences -------------------------------------------
    # A fence opened and never closed silently swallows the remainder of the
    # document, which is the single most destructive Markdown failure mode.
    $fenceLines = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*(```|~~~)') { $fenceLines += ($i + 1) }
    }
    if (($fenceLines.Count % 2) -ne 0) {
        Report $rel "unbalanced code fences ($($fenceLines.Count) markers - one is never closed)" $fenceLines[-1]
    }

    # Track which lines are inside a fence so table/heading checks skip them.
    $inFence = New-Object bool[] ($lines.Count + 1)
    $open = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*(```|~~~)') { $open = -not $open; $inFence[$i + 1] = $true; continue }
        $inFence[$i + 1] = $open
    }

    # ---- 2. table structure -------------------------------------------------
    # A GFM table needs: header row, delimiter row (---), then body rows. All
    # must agree on column count, or GitHub renders the block as plain text.
    function Get-ColumnCount {
        param([string]$Row)
        $t = $Row.Trim()
        if ($t.StartsWith('|')) { $t = $t.Substring(1) }
        if ($t.EndsWith('|')) { $t = $t.Substring(0, $t.Length - 1) }
        # Count unescaped pipes; \| inside a cell is a literal, not a separator.
        $count = 1
        for ($k = 0; $k -lt $t.Length; $k++) {
            if ($t[$k] -eq '|' -and ($k -eq 0 -or $t[$k - 1] -ne '\')) { $count++ }
        }
        return $count
    }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($inFence[$i + 1]) { continue }
        $line = $lines[$i]
        if ($line -notmatch '^\s*\|') { continue }
        # This is a table row. Its header is the first row of the run.
        if ($i -gt 0 -and $lines[$i - 1] -match '^\s*\|' -and -not $inFence[$i]) { continue }

        $header = $line
        $headerCols = Get-ColumnCount $header
        $delimIdx = $i + 1
        if ($delimIdx -ge $lines.Count -or $lines[$delimIdx] -notmatch '^\s*\|[\s:\-\|]+\|?\s*$') {
            Report $rel 'table header is not followed by a delimiter row (--- | ---)' ($i + 1)
            continue
        }
        $delimCols = Get-ColumnCount $lines[$delimIdx]
        if ($delimCols -ne $headerCols) {
            Report $rel "table delimiter has $delimCols column(s) but the header has $headerCols" ($delimIdx + 1)
        }
        # Body rows.
        for ($b = $delimIdx + 1; $b -lt $lines.Count; $b++) {
            if ($inFence[$b + 1]) { break }
            if ($lines[$b] -notmatch '^\s*\|') { break }
            if ($lines[$b] -match '^\s*\|[\s:\-\|]+\|?\s*$') {
                Report $rel 'stray delimiter row inside the table body' ($b + 1)
                continue
            }
            $bodyCols = Get-ColumnCount $lines[$b]
            if ($bodyCols -ne $headerCols) {
                Report $rel "table row has $bodyCols column(s), expected $headerCols" ($b + 1)
            }
        }
    }

    # ---- 3. headings: spacing and duplicates --------------------------------
    $seen = @{}
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($inFence[$i + 1]) { continue }
        if ($lines[$i] -notmatch '^(#{1,6})\s+(.+?)\s*$') { continue }
        $level = $Matches[1].Length
        $text = $Matches[2].Trim()

        # A heading needs a blank line before it, or several parsers treat it as
        # part of the preceding paragraph.
        if ($i -gt 0 -and $lines[$i - 1].Trim() -ne '' -and $lines[$i - 1] -notmatch '^---\s*$') {
            Report $rel "heading '$text' has no blank line before it" ($i + 1)
        }

        $key = "$level|$text"
        if ($seen.ContainsKey($key)) {
            Report $rel "duplicate heading '$text' at level $level (first at line $($seen[$key])) - anchors will collide" ($i + 1)
        }
        else { $seen[$key] = ($i + 1) }
    }

    # ---- 4. link hygiene ----------------------------------------------------
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($inFence[$i + 1]) { continue }
        $line = $lines[$i]
        if ($line -match '\]\(\s*\)') { Report $rel 'empty link target `[]()`' ($i + 1) }
        if ($line -match '\]\((TODO|TBD|FIXME|xxx|#\))') { Report $rel 'placeholder link target' ($i + 1) }
        if ($line -match '\[\s*\]\[') { Report $rel 'empty reference-style link text' ($i + 1) }
    }

    # ---- 5. no leftover edit markers ----------------------------------------
    # Must skip fenced regions: this project's documentation embeds real console
    # output, which is full of '================' banners. Without the fence check
    # every sample transcript is reported as a merge conflict.
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($inFence[$i + 1]) { continue }
        if ($lines[$i] -match '^\s*(<{7}|>{7}|={7})(\s|$)') {
            Report $rel 'leftover merge/conflict marker' ($i + 1)
        }
    }

    if ($fail -eq $fileFails) { $pass++; if (-not $Quiet) { Write-Host ("  OK    {0}" -f $rel) -ForegroundColor Green } }
}

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  DOC LINT: {0} file(s) checked, {1} clean, {2} issue(s)" -f $filesChecked, $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($fail -eq 0) { 0 } else { 1 })
