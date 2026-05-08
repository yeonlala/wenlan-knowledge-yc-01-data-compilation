@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
if not exist "clean_workspace.py" (
    echo ERROR: clean_workspace.py not found.
    pause & exit /b 1
)
where py >nul 2>&1
if not errorlevel 1 (
    py -3 "clean_workspace.py" --mock --yes
    goto :done
)
where python >nul 2>&1
if not errorlevel 1 (
    python "clean_workspace.py" --mock --yes
    goto :done
)
echo ERROR: Python not found.
pause
exit /b 1
:done
if errorlevel 1 pause & exit /b 1
echo Done.
pause
