@echo off
setlocal

set "RESOURCES_SRC=m:\06b. VH Project Engineering\Projecten\Engineering\Beheer VH\scripts\C# plugins\VH_addins\src\VH_Addin\Resources"
set "BIN_SRC_ROOT=m:\06b. VH Project Engineering\Projecten\Engineering\Beheer VH\scripts\C# plugins\VH_addins\src\VH_Addin\bin"
set "ADDIN_SRC=m:\06b. VH Project Engineering\Projecten\Engineering\Beheer VH\scripts\C# plugins\VH_addins\src\VH_Addin\VH_Tools.addin"
set "ADDINS_BASE=C:\Users\BasAarntzen-VHEngine\AppData\Roaming\Autodesk\Revit\Addins"

:: Deploy to 2023
echo Deploying to 2023...
if not exist "%ADDINS_BASE%\2023\Resources" mkdir "%ADDINS_BASE%\2023\Resources"
copy /Y "%BIN_SRC_ROOT%\Debug2023\VH_Tools.dll" "%ADDINS_BASE%\2023\"
copy /Y "%BIN_SRC_ROOT%\Debug2023\VH_Tools.pdb" "%ADDINS_BASE%\2023\"
copy /Y "%ADDIN_SRC%" "%ADDINS_BASE%\2023\"
copy /Y "%RESOURCES_SRC%\VH_Icon32.png" "%ADDINS_BASE%\2023\Resources\"

:: Deploy to 2024
echo Deploying to 2024...
if not exist "%ADDINS_BASE%\2024\Resources" mkdir "%ADDINS_BASE%\2024\Resources"
copy /Y "%BIN_SRC_ROOT%\Debug2024\VH_Tools.dll" "%ADDINS_BASE%\2024\"
copy /Y "%BIN_SRC_ROOT%\Debug2024\VH_Tools.pdb" "%ADDINS_BASE%\2024\"
copy /Y "%ADDIN_SRC%" "%ADDINS_BASE%\2024\"
copy /Y "%RESOURCES_SRC%\VH_Icon32.png" "%ADDINS_BASE%\2024\Resources\"

:: Deploy to 2025
echo Deploying to 2025...
if not exist "%ADDINS_BASE%\2025\Resources" mkdir "%ADDINS_BASE%\2025\Resources"
copy /Y "%BIN_SRC_ROOT%\Debug2025\VH_Tools.dll" "%ADDINS_BASE%\2025\"
copy /Y "%BIN_SRC_ROOT%\Debug2025\VH_Tools.pdb" "%ADDINS_BASE%\2025\"
copy /Y "%ADDIN_SRC%" "%ADDINS_BASE%\2025\"
copy /Y "%RESOURCES_SRC%\VH_Icon32.png" "%ADDINS_BASE%\2025\Resources\"

echo Deployment complete!
