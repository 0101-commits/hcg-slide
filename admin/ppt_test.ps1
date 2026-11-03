# PowerPoint 실제 시험 — 「HCG 장표」 단추를 눌러 작업창을 열고, 작업창을 원격 조작할 수 있게 디버그 포트를 켬
# 사용: powershell -NoProfile -ExecutionPolicy Bypass -File ppt_test.ps1 -Doc <시험 pptx>
# 열린 PowerPoint 가 있으면 건드리지 않고 멈춤(사용자 작업 보호).
param([Parameter(Mandatory = $true)][string]$Doc, [int]$Port = 9339)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Output 'STOP: PowerPoint 가 이미 켜져 있음'; exit 2 }

$exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE' -ErrorAction SilentlyContinue).'(default)'
if (-not $exe) { $exe = 'POWERPNT.EXE' }
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Port"
$p = Start-Process -FilePath $exe -ArgumentList "`"$Doc`"" -PassThru
Write-Output "PowerPoint pid $($p.Id)"

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::RootElement
$win = $null
for ($i = 0; $i -lt 90 -and -not $win; $i++) {
    Start-Sleep -Milliseconds 1000
    $c = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
    # 여는 중(시작 화면) 창이 아니라 문서 창이 뜰 때까지 기다림
    $win = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $c) | Where-Object { $_.Current.Name -like '*- PowerPoint' } | Select-Object -First 1
}
if (-not $win) { Write-Output 'FAIL: PowerPoint 창을 못 찾음'; exit 1 }
Write-Output "창: $($win.Current.Name)"

$btn = $null
for ($i = 0; $i -lt 40 -and -not $btn; $i++) {
    Start-Sleep -Milliseconds 1000
    $c = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, 'HCG 장표')
    $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $c)
    foreach ($e in $all) { if ($e.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button) { $btn = $e; break } }
}
if (-not $btn) { Write-Output 'FAIL: 리본에 「HCG 장표」 단추가 없음'; exit 1 }
Write-Output 'OK: 리본에 「HCG 장표」 단추 있음'
$inv = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
$inv.Invoke()
Write-Output 'OK: 단추 누름 — 작업창 여는 중'
