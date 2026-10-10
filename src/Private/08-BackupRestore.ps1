# ============================================================================
#  Backup / restore / retention  (finding A3: v3 deleted files with no undo)
# ----------------------------------------------------------------------------
#  A backup bundle is a ZIP containing:
#     backup-manifest.json    who/what/when, file list, job metadata, notes
#     F00012.SHD, 00012.SPL   the spool files exactly as they were
#
#  IMPORTANT, and stated plainly in the docs: copying spool files back into
#  PRINTERS\ does NOT re-queue the job. The Spooler builds its job table in
#  memory and only reads .SHD/.SPL for jobs it created in the current boot. A
#  restore is therefore a FORENSIC operation - it recovers the document payload
#  so an operator can re-submit it from the application - not a transactional
#  rollback. Claiming otherwise would be a lie in someone's runbook.
# ============================================================================

function Get-SpoolBackupInfo {
    <#
    .SYNOPSIS
        Reads backup-manifest.json out of a backup ZIP without extracting it.
    .OUTPUTS
        PSCustomObject { Path, Valid, Tool, ToolVersion, CreatedUtc, ComputerName,
                         SpoolDirectory, FileCount, Files, Jobs, RestoreNote,
                         Hash, SizeBytes, Error }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Path)

    $info = [ordered]@{
        Path           = $Path
        Valid          = $false
        Tool           = $null
        ToolVersion    = $null
        CreatedUtc     = $null
        ComputerName   = $null
        SpoolDirectory = $null
        FileCount      = 0
        Files          = @()
        Jobs           = @()
        RestoreNote    = $null
        Hash           = $null
        SizeBytes      = [long]0
        Error          = $null
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $info.Error = 'No backup path was supplied.'
        return [pscustomobject]$info
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $info.Error = "Backup file not found: $Path"
        return [pscustomobject]$info
    }

    $zip = $null
    try {
        $info.SizeBytes = (Get-Item -LiteralPath $Path -Force).Length
        $info.Hash = Get-FileHashSafe -Path $Path -Algorithm 'SHA256'

        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Path).Path)

        $entry = @($zip.Entries | Where-Object { $_.FullName -like '*backup-manifest.json' } | Select-Object -First 1)
        if ($entry.Count -eq 0 -or -not $entry[0]) {
            $info.Error = 'No backup-manifest.json inside the archive.'
            return [pscustomobject]$info
        }

        $reader = New-Object System.IO.StreamReader($entry[0].Open(), [System.Text.Encoding]::UTF8)
        try { $json = $reader.ReadToEnd() }
        finally { $reader.Dispose() }

        $manifest = $json | ConvertFrom-Json -ErrorAction Stop
        $info.Tool           = $manifest.Tool
        $info.ToolVersion    = $manifest.ToolVersion
        $info.CreatedUtc     = $manifest.CreatedUtc
        $info.ComputerName   = $manifest.ComputerName
        $info.SpoolDirectory = $manifest.SpoolDirectory
        $info.FileCount      = [int]$manifest.FileCount
        $info.Files          = @($manifest.Files)
        $info.Jobs           = @($manifest.Jobs)
        $info.RestoreNote    = $manifest.RestoreNote
        $info.Valid          = $true
    }
    catch {
        $info.Error = $_.Exception.Message
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
    return [pscustomobject]$info
}

function Expand-SpoolBackupArchive {
    <#
    .SYNOPSIS
        Extracts a backup ZIP into a destination directory.
    .DESCRIPTION
        Entry names are reduced to their file-name component before being joined
        to the destination, which is the standard defence against a "zip slip"
        archive that carries ..\..\ path segments.
    .OUTPUTS
        Number of files extracted.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ArchivePath,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Destination,
        [switch]$Overwrite
    )

    $count = 0
    if (-not (Test-Path -LiteralPath $Destination)) {
        New-Item -ItemType Directory -Path $Destination -Force -ErrorAction Stop | Out-Null
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $ArchivePath).Path)
        foreach ($entry in $zip.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.Name)) { continue }   # directory entry
            $safeName = [System.IO.Path]::GetFileName($entry.Name)
            if ([string]::IsNullOrWhiteSpace($safeName)) { continue }
            if ($safeName -eq '.' -or $safeName -eq '..') { continue }

            $target = Join-Path $Destination $safeName
            if ((Test-Path -LiteralPath $target) -and -not $Overwrite) { continue }
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, [bool]$Overwrite)
            $count++
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
    return $count
}

function Restore-SpoolerBackupInternal {
    <#
    .SYNOPSIS
        Implementation behind the public Restore-SpoolerBackup cmdlet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Path,
        [AllowNull()][AllowEmptyString()][string]$DestinationDirectory,
        [AllowNull()][AllowEmptyString()][string]$ExpectedHash,
        [switch]$Overwrite,
        [switch]$Force
    )

    $info = Get-SpoolBackupInfo -Path $Path
    if (-not $info.Valid) { throw "Backup '$Path' could not be read: $($info.Error)" }

    if ($ExpectedHash -and -not (Test-HashMatch -Expected $ExpectedHash -Actual $info.Hash)) {
        throw ("Backup integrity check failed.`n  expected: {0}`n  actual  : {1}" -f $ExpectedHash, $info.Hash)
    }

    if (-not [string]::IsNullOrWhiteSpace($DestinationDirectory)) {
        $dest = Resolve-DirectoryPath -Path $DestinationDirectory -Create
    }
    else {
        $stamp = Get-SafeFileNameStamp
        $parent = Split-Path -Parent (Resolve-Path -LiteralPath $Path).Path
        $name = "restored_{0}_{1}" -f (ConvertTo-SafeFileName -Value $info.ComputerName), $stamp
        $dest = Resolve-DirectoryPath -Path (Join-Path $parent $name) -Create
    }

    Write-LogInfo ("Restoring the backup of '{0}' taken {1} ({2} file(s), {3}) to '{4}'." -f `
        $info.ComputerName, $info.CreatedUtc, $info.FileCount, (Format-ByteSize $info.SizeBytes), $dest) -Component 'Restore'

    $extracted = Expand-SpoolBackupArchive -ArchivePath $Path -Destination $dest -Overwrite:$Overwrite

    return [pscustomobject]@{
        BackupPath        = $info.Path
        Valid             = $true
        ComputerName      = $info.ComputerName
        CreatedUtc        = $info.CreatedUtc
        ToolVersion       = $info.ToolVersion
        ManifestFileCount = $info.FileCount
        FilesRestored     = $extracted
        Destination       = $dest
        Sha256            = $info.Hash
        Jobs              = $info.Jobs
        Note              = 'Spool files cannot be re-queued by copying them back into PRINTERS. Re-submit the documents from the application, or hand the extracted payload to the user.'
    }
}

function Invoke-BackupRetention {
    <#
    .SYNOPSIS
        Deletes spool-backup_*.zip bundles older than the retention window.
    .NOTES
        Scoped strictly to the backup directory and the 'spool-backup_' prefix, so
        a mis-pointed -BackupPath can never become a recursive delete.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$BackupPath,
        [int]$KeepDays = $Script:DefaultLogRetentionDays
    )

    if ([string]::IsNullOrWhiteSpace($BackupPath) -or $KeepDays -le 0) { return 0 }
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Container)) { return 0 }

    $removed = 0
    $cutoff = (Get-Date).AddDays(-$KeepDays)
    try {
        $bundles = @(Get-ChildItem -LiteralPath $BackupPath -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like 'spool-backup_*.zip' -and $_.LastWriteTime -lt $cutoff })
        foreach ($b in $bundles) {
            try {
                Remove-Item -LiteralPath $b.FullName -Force -ErrorAction Stop
                $removed++
                Write-LogDebug "Backup retention removed '$($b.Name)' (older than $KeepDays days)." -Component 'Backup'
            }
            catch {
                Write-LogWarn "Backup retention could not remove '$($b.Name)': $($_.Exception.Message)" -Component 'Backup'
            }
        }
    }
    catch {
        # Retention is housekeeping; a failure must not affect the cleanup result.
    }
    return $removed
}

function Get-SpoolBackupInventory {
    <#
    .SYNOPSIS
        Lists every backup bundle in a directory with its metadata.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$BackupPath)

    if ([string]::IsNullOrWhiteSpace($BackupPath)) { return @() }
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Container)) { return @() }

    $out = New-Object System.Collections.Generic.List[object]
    $bundles = @(Get-ChildItem -LiteralPath $BackupPath -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -ieq '.zip' -and $_.Name -like 'spool-backup_*' } |
        Sort-Object LastWriteTime -Descending)

    foreach ($f in $bundles) {
        $info = Get-SpoolBackupInfo -Path $f.FullName
        $out.Add([pscustomobject]@{
                File         = $f.Name
                FullPath     = $f.FullName
                CreatedUtc   = $info.CreatedUtc
                ComputerName = $info.ComputerName
                FileCount    = $info.FileCount
                SizeBytes    = $f.Length
                Sha256       = $info.Hash
                Valid        = $info.Valid
                Error        = $info.Error
            })
    }
    return $out.ToArray()
}
