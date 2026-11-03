# 「HCG 장표」 설치 — 설치.bat 가 부름. 관리자 권한 없이 내 계정에만 등록함.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$dest = Join-Path $env:LOCALAPPDATA 'HCG'
$log = @()

function Say($t) { $script:log += $t }

try {
    # 1) 설명 파일(manifest) 복사
    $src = Join-Path $root 'addin\manifest.xml'
    [xml]$mx = Get-Content -LiteralPath $src -Encoding UTF8
    $id = $mx.OfficeApp.Id
    if (-not $id) { throw '설명 파일에서 추가 기능 ID를 못 찾음' }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -LiteralPath $src -Destination (Join-Path $dest 'manifest.xml') -Force

    # 2) 내 계정 레지스트리에 등록 → 다음 PowerPoint 실행 때 「HCG 장표」 단추가 뜸
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
    if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
    New-ItemProperty -Path $key -Name $id -Value (Join-Path $dest 'manifest.xml') -PropertyType String -Force | Out-Null
    Say "✓ 「HCG 장표」 추가 기능 등록 ($($mx.OfficeApp.Version))"

    # 3) Claude Code 용 스킬 복사
    $skSrc = Join-Path $root 'skills'
    if (Test-Path $skSrc) {
        $skDst = Join-Path $env:USERPROFILE '.claude\skills'
        New-Item -ItemType Directory -Force -Path $skDst | Out-Null
        $names = @()
        foreach ($d in Get-ChildItem -LiteralPath $skSrc -Directory) {
            $to = Join-Path $skDst $d.Name
            if (Test-Path $to) { Remove-Item -LiteralPath $to -Recurse -Force }
            Copy-Item -LiteralPath $d.FullName -Destination $to -Recurse -Force
            $names += $d.Name
        }
        Say "✓ Claude Code 스킬 복사: $($names -join ', ')"
    }

    # 4) 확인만 — 파이썬 · python-pptx(Claude Code 에서 스킬 스크립트를 돌릴 때만 필요)
    $py = $null
    foreach ($c in @('py', 'python')) { if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; break } }
    if ($py) {
        & $py -c "import pptx" 2>$null
        if ($LASTEXITCODE -eq 0) { Say '✓ 파이썬 · python-pptx 있음' }
        else { Say "· python-pptx 없음 — Claude Code 에서 스크립트를 쓰려면: $py -m pip install python-pptx" }
    } else {
        Say '· 파이썬 없음 — Claude 앱 · PowerPoint 에서는 상관없음(Claude Code 스크립트만 필요)'
    }

    # 5) PowerPoint 판 확인
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.VersionToReport) { Say "✓ Microsoft 365 PowerPoint $($c2r.VersionToReport)" }
    else { Say '! Microsoft 365(구독) PowerPoint 를 못 찾음 — 2016 · 2019 같은 영구 판에서는 안 됨' }

    $running = Get-Process POWERPNT -ErrorAction SilentlyContinue
    $tail = if ($running) { "`n지금 PowerPoint 가 켜져 있음 → 모두 닫고 다시 켜면 홈 리본에 「HCG 장표」 단추가 뜸." }
            else { "`nPowerPoint 를 켜면 홈 리본에 「HCG 장표」 단추가 뜸." }
    $tail += "`n`nClaude 앱 · PowerPoint 의 Claude 에서 스킬을 쓰려면 「Claude앱_스킬」 폴더의 zip 3개를`nClaude 설정 → 기능 → 스킬에서 한 번씩 올리세요(읽어주세요.txt)."
    $msg = ($log -join "`n") + "`n" + $tail
    Set-Content -LiteralPath (Join-Path $dest 'install.log') -Value ("$(Get-Date -Format s)`n" + $msg) -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show($msg, 'HCG 장표 — 설치 끝', 'OK', 'Information') | Out-Null
}
catch {
    $err = "설치하지 못함: $($_.Exception.Message)`n`n" + ($log -join "`n") + "`n`n이 창을 캡처해 관리자에게 보내 주세요."
    try { New-Item -ItemType Directory -Force -Path $dest | Out-Null; Set-Content -LiteralPath (Join-Path $dest 'install.log') -Value $err -Encoding UTF8 } catch {}
    [System.Windows.Forms.MessageBox]::Show($err, 'HCG 장표 — 설치 실패', 'OK', 'Error') | Out-Null
    exit 1
}
