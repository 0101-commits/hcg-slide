param([string]$Dir, [string]$Out)
# Windows 내장 OCR(한국어)로 폴더 안 PNG 마다 글을 읽어 JSON 으로 씀. 외부 연결 없음.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($task, $type) { $t = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($task)); $t.Wait(-1) | Out-Null; $t.Result }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('ko'))
$res = @{}
foreach ($f in Get-ChildItem -Path $Dir -Filter *.png) {
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($f.FullName)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $r = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  $res[$f.BaseName] = @{ text = $r.Text; lines = $r.Lines.Count }
  $stream.Dispose()
}
$res | ConvertTo-Json -Depth 3 | Set-Content -Path $Out -Encoding UTF8
"ocr=$($res.Count)"
