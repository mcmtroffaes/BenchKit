@echo off
setlocal enabledelayedexpansion
for /f "tokens=1* delims= " %%x in ("%*") do (
  echo Output: %%x
  echo Command: %%y
  start "parentconsole" /b cmd.exe /C "%%y >> %%x"
)
endlocal
