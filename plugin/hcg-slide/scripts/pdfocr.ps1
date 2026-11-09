param([string]$File, [string]$Out, [int]$Pages = 15)
# 글이 없는 PDF(스캔본)를 Windows 내장 기능으로 쪽마다 그림으로 그려 한국어 OCR. 외부 연결 없음. 결과 = 글(UTF-8) 파일.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$ext = [System.WindowsRuntimeSystemExtensions].GetMethods()
$asTaskGeneric = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$asTaskAction = ($ext | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function Await($task, $type) { $t = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($task)); $t.Wait(-1) | Out-Null; $t.Result }
function AwaitAction($task) { $t = $asTaskAction.Invoke($null, @($task)); $t.Wait(-1) | Out-Null }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('ko'))
$sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($File)) ([Windows.Storage.StorageFile])
$pdf = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
$n = [Math]::Min([int]$pdf.PageCount, $Pages)
$sb = New-Object System.Text.StringBuilder
for ($i = 0; $i -lt $n; $i++) {
  $page = $pdf.GetPage($i)
  $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $opt.DestinationWidth = 1800
  $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitAction ($page.RenderToStreamAsync($ms, $opt))
  $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ms)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $r = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  [void]$sb.AppendLine("[쪽 $($i + 1)]")
  foreach ($ln in $r.Lines) { [void]$sb.AppendLine($ln.Text) }
  $ms.Dispose(); $page.Dispose()
}
[IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
"pages=$n"
