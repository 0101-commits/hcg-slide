param([string]$File, [string]$Dir, [int]$W = 1920, [int]$H = 1080)
# 장 그림 내보내기: 창 없이 읽기 전용으로 열고, 우리 파일만 닫고, 시작 때 PowerPoint 가 꺼져 있었을 때만 끔
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Dir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$before = $app.Presentations.Count
try {
  $p = $app.Presentations.Open($File, -1, 0, 0)
  for ($i = 1; $i -le $p.Slides.Count; $i++) { $p.Slides.Item($i).Export((Join-Path $Dir ("R_{0:00}.png" -f $i)), 'PNG', $W, $H) }
  "slides=$($p.Slides.Count)"
  $p.Close()
} finally {
  if ($before -eq 0 -and $app.Presentations.Count -eq 0) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
