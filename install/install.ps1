# 「HCG 장표」 설치 — 설치.bat 가 부름. 관리자 권한 없이 내 계정에만 등록함.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$dest = Join-Path $env:LOCALAPPDATA 'HCG'
$catalogGuid = '{6B2D1C4E-8F3A-4E57-9C21-5A7E2B9D0C14}'  # 고정 — uninstall.ps1 과 같아야 함
$log = @()

function Say($t) { $script:log += $t }

try {
    # 1) 설명 파일(manifest) 복사
    $src = Join-Path $root 'addin\manifest.xml'
    [xml]$mx = Get-Content -LiteralPath $src -Encoding UTF8
    $id = $mx.OfficeApp.Id
    if (-not $id) { throw '설명 파일에서 추가 기능 ID를 못 찾음' }
    $cat = Join-Path $dest 'catalog'   # 이 폴더엔 설명 파일만 둠(Office 가 폴더 안 xml 을 모두 읽음)
    New-Item -ItemType Directory -Force -Path $cat | Out-Null
    Copy-Item -LiteralPath $src -Destination (Join-Path $cat 'manifest.xml') -Force
    $old = Join-Path $dest 'manifest.xml'  # 1.2.0 이전 자리
    if (Test-Path $old) { Remove-Item -LiteralPath $old -Force }

    # 2) 등록. 2026-09-25 Office 업데이트부터 「개발자 추가 기능」(Wef\Developer) 등록은 PowerPoint 를 끄면 잊힘
    #    (OfficeDev/office-js#6973). 신뢰 카탈로그(공유 폴더)로 한 번 추가한 것은 계속 남음.
    #    공유 폴더는 네트워크 경로여야 해서 드라이브 관리 공유(\\localhost\C$\…)로 내 폴더를 가리킴 — 새로 만들 것·관리자 권한 없음.
    $wef = 'HKCU:\Software\Microsoft\Office\16.0\Wef'
    $dev = "$wef\Developer"
    $catKey = "$wef\TrustedCatalogs\$catalogGuid"
    $share = $null
    if ($cat -match '^([A-Za-z]):\\(.*)$') { $share = "\\localhost\$($Matches[1].ToUpper())`$\$($Matches[2])" }
    elseif ($cat.StartsWith('\\')) { $share = $cat }
    $shareOk = $false
    if ($share) { try { $shareOk = Test-Path -LiteralPath (Join-Path $share 'manifest.xml') } catch { $shareOk = $false } }

    if ($shareOk) {
        if (-not (Test-Path $catKey)) { New-Item -Path $catKey -Force | Out-Null }
        New-ItemProperty -Path $catKey -Name 'Id' -Value $catalogGuid -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catKey -Name 'Url' -Value $share -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catKey -Name 'Flags' -Value 1 -PropertyType DWord -Force | Out-Null  # 1 = 메뉴에 보이기
        # 같은 ID 가 개발자 추가 기능에도 있으면 둘이 겹침 → 지움
        if ((Test-Path $dev) -and (Get-ItemProperty -Path $dev -Name $id -ErrorAction SilentlyContinue)) { Remove-ItemProperty -Path $dev -Name $id }
        $mode = 'catalog'
        Say "✓ 「HCG 장표」 추가 기능 등록 ($($mx.OfficeApp.Version)) — 공유 폴더 방식"
    } else {
        if (-not (Test-Path $dev)) { New-Item -Path $dev -Force | Out-Null }
        New-ItemProperty -Path $dev -Name $id -Value (Join-Path $cat 'manifest.xml') -PropertyType String -Force | Out-Null
        $mode = 'developer'
        Say "✓ 「HCG 장표」 추가 기능 등록 ($($mx.OfficeApp.Version)) — 개발자 추가 기능 방식"
        Say "! 이 PC 는 관리 공유($share)를 못 읽어 공유 폴더 방식을 못 씀(표준 계정이거나 관리 공유가 꺼짐)."
    }

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
    $tail = if ($running) { "`n지금 PowerPoint 가 켜져 있음 → 모두 닫고 다시 켠 뒤:" } else { "`nPowerPoint 를 켠 뒤:" }
    if ($mode -eq 'catalog') {
        $tail += "`n  처음 한 번만 — 홈 → 추가 기능 → 맨 아래 「추가 기능 더 보기」`n  → 위쪽 「공유 폴더」 → 「HCG 장표」 → 「추가」."
        $tail += "`n  그다음부터는 PowerPoint 를 켤 때마다 홈 리본에 「HCG 장표」 단추가 그대로 있음."
    } else {
        $tail += "`n  홈 → 추가 기능 → 「개발자 추가 기능」의 「HCG 장표」를 누르면 단추가 뜸."
        $tail += "`n  지금 Office 문제로 PowerPoint 를 켤 때마다 다시 눌러야 함."
    }
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
