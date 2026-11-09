param([string]$File, [string]$Out, [string]$Slides = "")
# -Slides "3,8,12" 이면 그 장만 잼(줄여 쓴 뒤 다시 잴 때)
# 겹침 검사용: 장마다 모든 도형(묶음 안 포함)의 자리와, 글이 실제로 차지하는 자리(BoundLeft/Top/Width/Height)를 JSON 으로 씀
# 규칙: 창 없이 읽기 전용으로 열기, 우리 파일만 닫기, 시작 때 PowerPoint 가 꺼져 있었을 때만 끔
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$before = $app.Presentations.Count
$res = New-Object System.Collections.ArrayList
function Add($sh, $si) {
  if ($sh.Type -eq 6) { foreach ($g in $sh.GroupItems) { Add $g $si }; return }
  $o = [ordered]@{ slide = $si; id = $sh.Id; type = $sh.Type; x = [math]::Round($sh.Left, 1); y = [math]::Round($sh.Top, 1); w = [math]::Round($sh.Width, 1); h = [math]::Round($sh.Height, 1); t = '' }
  if ($sh.HasTextFrame) {
    $tr = $sh.TextFrame.TextRange
    if ($tr.Text.Trim().Length -gt 0) {
      $o.t = $tr.Text.Substring(0, [math]::Min(30, $tr.Text.Length))
      $o.wrap = [int]$sh.TextFrame.WordWrap
      $o.bx = [math]::Round($tr.BoundLeft, 1); $o.by = [math]::Round($tr.BoundTop, 1); $o.bw = [math]::Round($tr.BoundWidth, 1); $o.bh = [math]::Round($tr.BoundHeight, 1)
    }
  }
  [void]$res.Add([pscustomobject]$o)
}
try {
  $p = $app.Presentations.Open($File, -1, 0, 0)
  $only = @($Slides -split ',' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
  for ($i = 1; $i -le $p.Slides.Count; $i++) {
    if ($only.Count -gt 0 -and -not ($only -contains $i)) { continue }
    foreach ($sh in $p.Slides.Item($i).Shapes) { Add $sh $i }
  }
  $p.Close()
} finally {
  if ($before -eq 0 -and $app.Presentations.Count -eq 0) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
$res | ConvertTo-Json -Depth 3 -Compress | Set-Content -Path $Out -Encoding UTF8
"shapes=$($res.Count)"
