@echo off
@cls
@echo Use VS to export your keyboard settings, then close it to continue
@echo.
pause
@"%VS100COMNTOOLS%\..\IDE\devenv.exe" /Command Tools.ImportandExportSettings
@echo.
@echo deleting old AddIn files
@echo.
@del "%USERPROFILE%\Documents\Visual Studio 2010\Addins\brief.addin"
@del "%USERPROFILE%\Documents\Visual Studio 2010\Addins\brief.dll"
@echo.
@echo.
@echo VS2010 will open (for no apparent reason), please close it to continue
@echo.
pause
@"%VS100COMNTOOLS%\..\IDE\devenv.exe"
@echo.
@copy .\BRIEF\BRIEF.AddIn "%USERPROFILE%\Documents\Visual Studio 2010\Addins"
@copy .\BRIEF\bin\BRIEF.dll "%USERPROFILE%\Documents\Visual Studio 2010\Addins"
@echo.
@echo.
@echo Use VS to import the keyboard settings you just exported, then close it to finish
@echo.
pause
@"%VS100COMNTOOLS%\..\IDE\devenv.exe" /ResetAddin BRIEF.Connect.BRIEFAltA