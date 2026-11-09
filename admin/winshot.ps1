# PowerPoint 창을 그대로 PNG 로 찍음(사용설명서용 실제 화면). 사용: powershell -NoProfile -File winshot.ps1 -Out a.png [-Activate]
param([Parameter(Mandatory = $true)][string]$Out, [switch]$Activate)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing, System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class W {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
'@
[W]::SetProcessDPIAware() | Out-Null
$p = Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if (-not $p) { Write-Output 'FAIL: PowerPoint 창 없음'; exit 1 }
$h = $p.MainWindowHandle
if ($Activate) { [W]::ShowWindow($h, 3) | Out-Null; [W]::SetForegroundWindow($h) | Out-Null; Start-Sleep -Milliseconds 700 }
$r = New-Object W+RECT
[W]::GetWindowRect($h, [ref]$r) | Out-Null
# 최대화 창은 화면 밖 테두리 몇 px 가 있음 → 화면 안으로 자름
$scr = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$L = [Math]::Max($r.L, 0); $T = [Math]::Max($r.T, 0)
$w = $r.R - $L; $hh = $r.B - $T
if ($scr) { $w = [Math]::Min($w, $scr.Width - $L); $hh = [Math]::Min($hh, $scr.Height - $T) }
$bmp = New-Object System.Drawing.Bitmap $w, $hh
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($L, $T, 0, 0, $bmp.Size)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "OK $Out ${w}x${hh}"
