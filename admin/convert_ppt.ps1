# 옛 .ppt 원본을 .pptx 로 바꿈(창고에 넣을 장표가 든 것만) — PowerPoint 자동 조작
# 사용: powershell -NoProfile -ExecutionPolicy Bypass -File convert_ppt.ps1
# 켜진 PowerPoint 에 다른 문서가 열려 있으면 멈춤. 원본(OneDrive)은 복사해서 씀.
$ErrorActionPreference = 'Stop'
$ids = @('040d1b1d6a', '88f72ec45a', '1e95fc1b8f', '0938f575ea', '208e7266ec', '4635dad850', '0984d549c4', '1ab5b64f56', '6c86b1e2b8')
$src = 'C:\Users\cgpar\OneDrive - 휴먼컨설팅그룹\08 Reference\01 제안서'
$work = 'C:\Users\cgpar\hcg-proposal\work\ref\conv'
New-Item -ItemType Directory -Force -Path $work | Out-Null
$files = Get-Content 'C:\Users\cgpar\hcg-proposal\work\ref\files.json' -Encoding UTF8 -Raw | ConvertFrom-Json

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Output 'STOP: PowerPoint 가 켜져 있음 — 닫고 다시'; exit 2 }
$app = New-Object -ComObject PowerPoint.Application
try {
    foreach ($id in $ids) {
        $f = $files | Where-Object { $_.id -eq $id } | Select-Object -First 1
        $out = Join-Path $work "$id.pptx"
        if (Test-Path $out) { Write-Output "건너뜀 $id"; continue }
        $tmp = Join-Path $work "$id.ppt"
        Copy-Item -LiteralPath (Join-Path $src $f.rel) -Destination $tmp -Force
        $p = $app.Presentations.Open($tmp, -1, 0, 0)   # 읽기 전용 · 창 없이
        $p.SaveAs($out, 24)                             # ppSaveAsOpenXMLPresentation
        $p.Close()
        Remove-Item -LiteralPath $tmp -Force
        Write-Output "변환 $id $($f.rel)"
    }
}
finally {
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
