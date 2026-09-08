@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM ABSA Converter Tool - apply/build helper for the v3 changes
REM Run this from the repository ROOT (the folder that contains
REM Absa.sln), AFTER copying the updated files into place:
REM   AbsaConverterTool\Helper\FileHelper.cs
REM   AbsaConverterTool\MainWindow.xaml.cs
REM   AbsaConverterTool\MainWindow.xaml
REM   AbsaConverterTool\AbsaConverterTool.csproj
REM This script does NOT create a new project or repo - it only
REM restores/builds the existing one in place.
REM ============================================================

set SLN=Absa.sln

if not exist "%SLN%" (
    echo [ERROR] %SLN% not found in the current directory.
    echo Run this script from the repository root, next to Absa.sln.
    exit /b 1
)

echo === Step 1: Restoring NuGet packages (ClosedXML, System.Data.OleDb, WindowsAPICodePack-Shell) ===
dotnet restore "%SLN%"
if errorlevel 1 (
    echo [WARN] dotnet restore reported an error - continuing, MSBuild restore in Step 3 will retry it.
)

echo.
echo === Step 2: Locating Visual Studio MSBuild ===
REM dotnet build CANNOT build this project - the ADOX COM reference requires
REM the full-framework MSBuild that ships with Visual Studio.
set MSBUILD=
set VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe

if not exist "%VSWHERE%" (
    echo [ERROR] vswhere.exe not found - Visual Studio 2022 does not appear to be installed.
    echo Install Visual Studio 2022 with the ".NET desktop development" workload, then re-run this script.
    exit /b 1
)

for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do (
    set MSBUILD=%%i
)

if "%MSBUILD%"=="" (
    echo [ERROR] Could not locate MSBuild.exe via vswhere.
    echo Install/repair Visual Studio 2022 with the ".NET desktop development" workload.
    exit /b 1
)

echo Found MSBuild at: %MSBUILD%

echo.
echo === Step 3: Building %SLN% (Debug^|x86) ===
"%MSBUILD%" "%SLN%" /p:Configuration=Debug /p:Platform=x86 /t:Restore,Build /m

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. Common causes on this project:
    echo   - ADOX COM reference not registered / Access Database Engine ^(32-bit^) not installed
    echo   - NuGet restore did not pick up the new ClosedXML package
    echo   - A stale bin\obj folder from a previous build ^(try: rmdir /s /q AbsaConverterTool\bin AbsaConverterTool\obj^)
    echo Review the MSBuild output above for the exact error and re-run this script.
    exit /b 1
)

echo.
echo === Build succeeded ===
echo.
echo Manual verification steps ^(cannot be scripted - requires the running UI^):
echo   1. Launch AbsaConverterTool\bin\x86\Debug\net8.0-windows\AbsaConverterTool.exe
echo   2. Confirm the title bar reads "ABSA Tool v03".
echo   3. Prepare a small test folder with sample .inp-style files covering:
echo        - a known product code (e.g. AGVR)
echo        - an unknown product code
echo        - a missing product code
echo        - AGVQ records (Signature Credit / Signature Debit)
echo        - chip data at, under, and over the 1020-char combined limit
echo        - card holder names shorter than, exactly, and longer than 26 chars
echo        - multiple products mixed in one input file
echo   4. Click "Choose Folder", review the preview/log and the failed-record count.
echo   5. Click "Process Files", then confirm:
echo        - one .mdb per product bucket, named ^<input^>_^<ProductName^>.mdb
echo        - AGVQ records land in a single .mdb, shown as "Signature Credit + Signature Debit"
echo        - Failures.xlsx is created in the same folder when any record failed, with columns:
echo          Masked PAN, File Name, Record Index, Product Code, Reason
echo   See the delivered chat summary / README.md for the full verification checklist.

endlocal
