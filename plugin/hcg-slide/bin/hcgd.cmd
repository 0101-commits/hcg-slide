@echo off
rem HCG slide Draft steps - runs scripts\hcgd.py with the bundled Python (keep this file ASCII: cmd.exe reads it in the OEM code page)
"%~dp0..\python\python.exe" -X utf8 "%~dp0..\scripts\hcgd.py" %*
