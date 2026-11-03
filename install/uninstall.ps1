# 「HCG 장표」 지우기 — 제거.bat 가 부름. 레지스트리 한 줄 · %LOCALAPPDATA%\HCG · HCG 스킬 폴더를 지움.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$dest = Join-Path $env:LOCALAPPDATA 'HCG'
$log = @()
try {
    $src = Join-Path $root 'addin\manifest.xml'
    $id = $null
    if (Test-Path $src) { [xml]$mx = Get-Content -LiteralPath $src -Encoding UTF8; $id = $mx.OfficeApp.Id }
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Wef\Developer'
    if ($id -and (Test-Path $key) -and (Get-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue)) {
        Remove-ItemProperty -Path $key -Name $id
        $log += '✓ 「HCG 장표」 등록을 지움'
    } else { $log += '· 등록이 없음' }
    if (Test-Path $dest) { Remove-Item -LiteralPath $dest -Recurse -Force; $log += "✓ $dest 지움" }
    $skDst = Join-Path $env:USERPROFILE '.claude\skills'
    $skSrc = Join-Path $root 'skills'
    if (Test-Path $skSrc) {
        foreach ($d in Get-ChildItem -LiteralPath $skSrc -Directory) {
            $to = Join-Path $skDst $d.Name
            if (Test-Path $to) { Remove-Item -LiteralPath $to -Recurse -Force; $log += "✓ 스킬 $($d.Name) 지움" }
        }
    }
    $log += "`nPowerPoint 를 다시 켜면 단추가 사라짐. Claude 앱에 올린 스킬은 Claude 설정 → 기능 → 스킬에서 지우세요."
    [System.Windows.Forms.MessageBox]::Show(($log -join "`n"), 'HCG 장표 — 지우기 끝', 'OK', 'Information') | Out-Null
}
catch {
    [System.Windows.Forms.MessageBox]::Show("지우지 못함: $($_.Exception.Message)", 'HCG 장표', 'OK', 'Error') | Out-Null
    exit 1
}
