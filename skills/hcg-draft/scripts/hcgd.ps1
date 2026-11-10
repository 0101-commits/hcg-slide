# HCG proposal Draft step runner (keep this file ASCII: Windows PowerShell 5.1 reads BOM-less files in the ANSI code page).
# First run: downloads the runtime (embedded Python + step scripts, see ..\runtime.json) from GitHub,
# checks its SHA-256 and unpacks it to %LOCALAPPDATA%\HCG\draft\runtime\<version>. Later runs start at once.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File "<skill>\scripts\hcgd.ps1" <step> [work folder] [more]
#   workdir <short name>  -> makes <Documents>\HCG Draft\<short>_<YYMMDD>[_2..] with an empty materials (U+C790 U+B8CC) folder, prints {"dir": ...}
#   any other step        -> runs scripts\hcgd.py <step> <data folder> [work folder] [more] (one JSON line)
# Own failures print {"error_code": "..."}; SKILL.md says what to tell the user for each code.
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

function Out-Json($obj) { [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress)) }
function Fail($code, $detail) { Out-Json @{ error_code = $code; detail = "$detail" }; exit 1 }

if (-not $Rest -or $Rest.Count -lt 1) { Fail 'no_step' '' }
$step = $Rest[0]
$more = @($Rest | Select-Object -Skip 1)

if ($step -eq 'workdir') {
    if ($more.Count -lt 1 -or -not $more[0].Trim()) { Fail 'no_name' '' }
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $short = -join ($more[0].Trim().ToCharArray() | Where-Object { $bad -notcontains $_ })
    $docs = [Environment]::GetFolderPath('MyDocuments')
    $parent = Join-Path $docs 'HCG Draft'
    $name = $short + '_' + (Get-Date -Format 'yyMMdd')
    $dir = Join-Path $parent $name
    $n = 2
    while (Test-Path -LiteralPath $dir) { $dir = Join-Path $parent ($name + '_' + $n); $n++ }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $sub = [string][char]0xC790 + [char]0xB8CC   # materials folder name
    New-Item -ItemType Directory -Force -Path (Join-Path $dir $sub) | Out-Null
    Out-Json @{ dir = $dir }
    exit 0
}

$skill = Split-Path -Parent $PSScriptRoot
try { $info = [IO.File]::ReadAllText((Join-Path $skill 'runtime.json'), $utf8) | ConvertFrom-Json }
catch { Fail 'no_runtime_json' $_.Exception.Message }
if (-not $env:LOCALAPPDATA) { Fail 'not_windows' '' }
$base = Join-Path $env:LOCALAPPDATA 'HCG\draft'
$data = Join-Path $base 'data'
$rtRoot = Join-Path $base 'runtime'
$rt = Join-Path $rtRoot $info.version
$py = Join-Path $rt 'python\python.exe'

if (-not (Test-Path -LiteralPath (Join-Path $rt 'ready.txt'))) {
    New-Item -ItemType Directory -Force -Path $rtRoot, $data | Out-Null
    $zip = Join-Path $base ('runtime-' + $info.version + '.zip')
    $got = $false; $why = ''
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -UseBasicParsing -Uri $info.url -OutFile $zip -TimeoutSec 300
        $got = $true
    } catch { $why = $_.Exception.Message }
    if (-not $got -and (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
        & curl.exe -sSfL --max-time 300 -o $zip $info.url 2>$null
        if ($LASTEXITCODE -eq 0) { $got = $true } else { $why += " / curl $LASTEXITCODE" }
    }
    if (-not $got) { Fail 'download' $why }
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zip).Hash.ToLower()
    if ($hash -ne ([string]$info.sha256).ToLower()) { Remove-Item -LiteralPath $zip -Force; Fail 'hash' $hash }
    $tmp = $rt + '.tmp'
    try {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($zip, $tmp)
        if (Test-Path -LiteralPath $rt) { Remove-Item -LiteralPath $rt -Recurse -Force }
        Rename-Item -LiteralPath $tmp -NewName (Split-Path -Leaf $rt)
        [IO.File]::WriteAllText((Join-Path $rt 'ready.txt'), [string]$info.version)
    } catch { Fail 'unpack' $_.Exception.Message }
    Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $rtRoot -Directory | Where-Object { $_.Name -ne $info.version } |
        ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Recurse -Force } catch {} }
}
if (-not (Test-Path -LiteralPath $py)) { Fail 'unpack' 'python.exe missing' }

# A PC that used the old hcg-slide plugin already has a locked key (same Windows user, DPAPI) - reuse it once.
$keyNew = Join-Path $data 'key.dat'
$keyOld = Join-Path $env:USERPROFILE '.claude\plugins\data\hcg-slide-hcg\key.dat'
if (-not (Test-Path -LiteralPath $keyNew) -and (Test-Path -LiteralPath $keyOld)) {
    New-Item -ItemType Directory -Force -Path $data | Out-Null
    Copy-Item -LiteralPath $keyOld -Destination $keyNew
}

$argv = @($step, $data) + $more
& $py -X utf8 (Join-Path $rt 'scripts\hcgd.py') @argv
exit $LASTEXITCODE
