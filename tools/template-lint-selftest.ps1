$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$lint = Join-Path $root 'tools/Invoke-TemplateLint.ps1'
$exe = (Get-Process -Id $PID).Path

$pass = 0; $fail = 0
function Expect {
    param([string]$Label, [scriptblock]$Build, [int]$WantExit, [string]$WantText = '')
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psqtpl-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path (Join-Path $tmp '.github/ISSUE_TEMPLATE') -Force
    try {
        & $Build $tmp
        $out = Join-Path ([IO.Path]::GetTempPath()) ("psqtpl-o-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        # -Root, NOT -WorkingDirectory: the linter resolves the tree from its own
        # location by default, so a changed working directory would be ignored and
        # every fixture would appear to "pass" for the wrong reason. That is
        # exactly how this self-test first came back 8/8 green on broken input.
        $proc = Start-Process -FilePath $exe `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $lint, '-Root', $tmp, '-Quiet') `
            -NoNewWindow -Wait -PassThru -RedirectStandardOutput $out
        $text = if (Test-Path -LiteralPath $out) { Get-Content -LiteralPath $out -Raw } else { '' }
        Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue

        $ok = ($proc.ExitCode -eq $WantExit)
        if ($ok -and $WantText) { $ok = ($text -match [regex]::Escape($WantText)) }
        if ($ok) { $script:pass++; Write-Host ("  FIRES/QUIET OK  {0}" -f $Label) -ForegroundColor Green }
        else {
            $script:fail++
            Write-Host ("  WRONG           {0}" -f $Label) -ForegroundColor Red
            Write-Host ("                  wanted exit {0}{1}, got {2}" -f $WantExit,
                $(if ($WantText) { " and text '$WantText'" } else { '' }), $proc.ExitCode) -ForegroundColor DarkRed
            if ($text) { Write-Host ("                  output: {0}" -f ($text -replace '\s+', ' ')) -ForegroundColor DarkYellow }
        }
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host '=== must FIRE (exit 1) on real defects ===' -ForegroundColor Cyan

Expect -Label 'invalid escape \C in a double-quoted scalar' -WantExit 1 -WantText 'invalid escape' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/config.yml') -Encoding UTF8 -Value @'
blank_issues_enabled: false
contact_links:
  - name: Ref
    url: https://example.com
    about: "Run .\Clear-Thing.ps1 -Help now."
'@
}

Expect -Label 'markdown element using content instead of value' -WantExit 1 -WantText "uses 'content'" -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug
description: Something broke.
body:
  - type: markdown
    attributes:
      content: |
        Hello there.
'@
}

Expect -Label 'markdown element with no text attribute at all' -WantExit 1 -WantText 'has no attributes.value' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug
description: Something broke.
body:
  - type: markdown
    attributes:
      render: text
'@
}

Expect -Label 'textarea element with no label' -WantExit 1 -WantText 'has no attributes.label' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug
description: Something broke.
body:
  - type: textarea
    attributes:
      placeholder: tell us more
'@
}

Expect -Label 'dropdown with no options' -WantExit 1 -WantText 'has no options' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug
description: Something broke.
body:
  - type: dropdown
    attributes:
      label: Version
'@
}

Expect -Label 'unknown element type' -WantExit 1 -WantText 'unknown type' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug
description: Something broke.
body:
  - type: richtext
    attributes:
      label: Details
'@
}

Expect -Label 'form template missing name/description/body' -WantExit 1 -WantText 'required name/description/body' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
title: Bug
body:
  - type: markdown
    attributes:
      value: hi
'@
}

Expect -Label 'contact link missing about' -WantExit 1 -WantText 'contact link has name/url/about' -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/config.yml') -Encoding UTF8 -Value @'
blank_issues_enabled: false
contact_links:
  - name: Ref
    url: https://example.com
'@
}

Write-Host ''
Write-Host '=== must stay QUIET (exit 0) on legitimate patterns ===' -ForegroundColor Cyan

Expect -Label 'valid form: markdown value + labelled textarea + dropdown with options' -WantExit 0 -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/bug_report.yml') -Encoding UTF8 -Value @'
name: Bug report
description: Something did not behave as documented.
title: "[Bug] "
labels: ["bug"]
body:
  - type: markdown
    attributes:
      value: |
        Thanks for reporting this.
  - type: textarea
    id: repro
    attributes:
      label: Steps to reproduce
      placeholder: tell us more
    validations:
      required: true
  - type: dropdown
    id: version
    attributes:
      label: Version
      options:
        - 4.0.0
        - 3.1.0
'@
}

Expect -Label 'backslashes inside a run: | block scalar are NOT escapes' -WantExit 0 -Build {
    param($d)
    $null = New-Item -ItemType Directory -Path (Join-Path $d '.github/workflows') -Force
    Set-Content -LiteralPath (Join-Path $d '.github/workflows/ci.yml') -Encoding UTF8 -Value @'
name: CI
jobs:
  gate:
    runs-on: windows-latest
    steps:
      - name: Build
        shell: pwsh
        run: |
          $raw = Get-Content ./src/x.ps1 -Raw
          $ver = [regex]::Match($raw, "\`$Script:ScriptVersion\s+=\s+'([^']+)'").Groups[1].Value
          if (-not $ver) { throw 'no version' }
          $p = 'C:\Windows\System32\spool\PRINTERS'
'@
}

Expect -Label 'single-quoted scalar with a Windows path' -WantExit 0 -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/config.yml') -Encoding UTF8 -Value @'
blank_issues_enabled: false
contact_links:
  - name: Ref
    url: https://example.com
    about: 'Run .\Clear-Thing.ps1 -Help for the reference.'
'@
}

Expect -Label 'correctly escaped double-quoted scalar' -WantExit 0 -Build {
    param($d)
    Set-Content -LiteralPath (Join-Path $d '.github/ISSUE_TEMPLATE/config.yml') -Encoding UTF8 -Value @'
blank_issues_enabled: false
contact_links:
  - name: Ref
    url: https://example.com
    about: "Run .\\Clear-Thing.ps1 -Help for the reference."
'@
}

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  TEMPLATE-LINT SELF-TEST: {0} passed, {1} failed" -f $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($fail -eq 0) { 0 } else { 1 })
