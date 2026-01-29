@echo off
echo Building VH Daglichttool Plugin...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" VH_DaglichtPlugin.csproj /p:Configuration=Debug /t:Rebuild /v:minimal
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Plugin deployed to:
    echo %APPDATA%\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\
    echo.
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
)
pause
