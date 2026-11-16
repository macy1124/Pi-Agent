#Requires -Version 5.1
<#
  Install anthropics/skills :: skills/pptx into a project's .pi/skills/pptx

  - Downloads the exact pinned commit (codeload zipball), no git required.
  - Copies the COMPLETE skill directory (SKILL.md, LICENSE.txt, scripts/, schemas/).
  - Records the source commit + per-file SHA256 into .source.json.
  - Does NOT install any dependencies (no pip/npm/npx).
  - If the target already exists it is MOVED ASIDE (never deleted), so the
    previous copy stays recoverable.
#>
[CmdletBinding()]
param(
    [string]$Commit = '1780c5e5a552fd4d1c9d6a7b0cab26b0c2bef064',
    [string]$Repo     = 'anthropics/skills',
    [string]$SkillPath = 'skills/pptx',
    [string]$Dest     = 'E:\workspace\pi-skill\.pi\skills\pptx'
)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

function Show-Tree {
    param([string]$Root, [string]$Prefix = '')
    $items = Get-ChildItem -LiteralPath $Root -Force | Sort-Object `
        @{ Expression = { -not $_.PSIsContainer } }, Name
    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items[$i]
        $last = ($i -eq $items.Count - 1)
        $branch = if ($last) { '\-- ' } else { '|-- ' }
        if ($item.PSIsContainer) {
            $n = (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force | Measure-Object).Count
            Write-Output ('{0}{1}{2}/   ({3} files)' -f $Prefix, $branch, $item.Name, $n)
            $next = $Prefix + $(if ($last) { '    ' } else { '|   ' })
            Show-Tree -Root $item.FullName -Prefix $next
        }
        else {
            Write-Output ('{0}{1}{2}   ({3:N0} B)' -f $Prefix, $branch, $item.Name, $item.Length)
        }
    }
}

$zip  = Join-Path $env:TEMP ("$($Repo -replace '/', '-')-$Commit.zip")
$work = Join-Path $env:TEMP ("pi-skill-$Commit")

Write-Host "== source =="
Write-Host "  repo   : $Repo"
Write-Host "  path   : $SkillPath"
Write-Host "  commit : $Commit"
Write-Host "  dest   : $Dest"
Write-Host ""

Write-Host "== download =="
$url = "https://codeload.github.com/$Repo/zip/$Commit"
Write-Host "  $url"
Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
Write-Host ("  -> {0:N0} bytes" -f (Get-Item -LiteralPath $zip).Length)
Write-Host ""

Write-Host "== extract =="
if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
Expand-Archive -LiteralPath $zip -DestinationPath $work -Force

$src = $null
foreach ($top in (Get-ChildItem -LiteralPath $work -Directory -Force)) {
    $candidate = Join-Path $top.FullName ($SkillPath -replace '/', '\')
    if (Test-Path -LiteralPath (Join-Path $candidate 'SKILL.md')) { $src = $candidate; break }
}
if (-not $src) {
    throw "Unexpected archive layout: no '$SkillPath' with a SKILL.md under $work"
}
Write-Host "  $src"
Write-Host ""

# ---- sanity: the pieces the skill needs at runtime ------------------------
$required = @(
    'SKILL.md',
    'LICENSE.txt',
    'scripts\add_slide.py',
    'scripts\clean.py',
    'scripts\thumbnail.py',
    'scripts\apply_theme.js',
    'scripts\slide_text.py',
    'scripts\office\validate.py',
    'scripts\office\soffice.py',
    'scripts\office\validators\pptx.py',
    'scripts\office\validators\base.py',
    'scripts\office\helpers\__init__.py',
    'scripts\office\schemas\ISO-IEC29500-4_2016\pml.xsd',
    'scripts\office\schemas\mce\mc.xsd'
)
$missing = @()
foreach ($r in $required) {
    if (-not (Test-Path -LiteralPath (Join-Path $src $r))) { $missing += $r }
}
if ($missing.Count -gt 0) {
    throw ("Upstream at this commit is missing expected files:`n  " + ($missing -join "`n  "))
}
Write-Host "== sanity check: all $($required.Count) expected files present =="
Write-Host ""

# ---- install -------------------------------------------------------------
if (Test-Path -LiteralPath $Dest) {
    $bak = "$Dest.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Write-Host "== target exists -> moving aside =="
    Write-Host "  $Dest"
    Write-Host "  -> $bak"
    Move-Item -LiteralPath $Dest -Destination $bak
}
New-Item -ItemType Directory -Force -Path $Dest | Out-Null

Write-Host "== copy =="
Copy-Item -Path (Join-Path $src '*') -Destination $Dest -Recurse -Force
Write-Host "  done"
Write-Host ""

# ---- provenance ----------------------------------------------------------
Write-Host "== record provenance =="
$files = Get-ChildItem -LiteralPath $Dest -Recurse -File -Force | Sort-Object FullName
$map = [ordered]@{}
foreach ($f in $files) {
    $rel = $f.FullName.Substring($Dest.Length).TrimStart('\') -replace '\\', '/'
    $map[$rel] = [ordered]@{
        sha256 = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash.ToLower()
        size   = $f.Length
    }
}

$manifest = [ordered]@{
    repo        = $Repo
    path        = $SkillPath
    ref         = 'main'
    commit      = $Commit
    commitUrl   = "https://github.com/$Repo/commit/$Commit"
    commitDate  = '2026-11-14T15:16:42Z'
    commitTitle = 'Update docx, pptx, xlsx, and pdf skills (#2217)'
    retrievedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    via         = 'codeload zipball'
    fileCount   = $map.Count
    note        = 'Dependencies were NOT installed.'
    files       = $map
}
$manifestPath = Join-Path $Dest '.source.json'
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Host ("  {0}  ({1} files hashed)" -f $manifestPath, $map.Count)
Write-Host ""

# ---- report --------------------------------------------------------------
Write-Host "== installed tree =="
Write-Output ('{0}\' -f $Dest)
Show-Tree -Root $Dest
Write-Host ""
Write-Host "== summary =="
Write-Host ("  files      : {0}" -f $map.Count)
Write-Host ("  bytes      : {0:N0}" -f (($files | Measure-Object -Property Length -Sum).Sum))
Write-Host ("  commit     : {0}" -f $Commit)
Write-Host ("  zip left at: {0}" -f $zip)
