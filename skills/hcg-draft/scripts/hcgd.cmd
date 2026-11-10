@echo off
rem HCG proposal Draft step runner - calls hcgd.ps1 next to this file (keep ASCII)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0hcgd.ps1" %*
