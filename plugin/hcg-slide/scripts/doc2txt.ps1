param([string]$File, [string]$Out)
# 옛 Word(.doc) 파일을 Word 로 읽기 전용으로 열어 글만 뽑음. 원본은 바꾸지 않음. Word 가 없으면 실패.
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.AutomationSecurity = 3  # 문서 매크로 끔(msoAutomationSecurityForceDisable)
try {
  $d = $w.Documents.Open($File, $false, $true, $false)
  $t = $d.Content.Text
  $d.Close($false)
} finally {
  if ($w.Documents.Count -eq 0) { $w.Quit() }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
[IO.File]::WriteAllText($Out, ($t -replace "`r", "`n"), (New-Object Text.UTF8Encoding($false)))
"chars=$($t.Length)"
