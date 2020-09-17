@echo off
@set root=%~dp0
if %PROCESSOR_ARCHITECTURE% == AMD64 (
if exist "%windir%\Syswow64\node.dll" goto ends
xcopy "%root%node.dll" "%windir%\Syswow64\" /s/y
regsvr32 "%windir%\Syswow64\node.dll" /s
) else (
if exist "%windir%\System32\node.dll" goto ends
xcopy "%root%node.dll" "%windir%\System32\" /s/y
regsvr32 "%windir%\System32\node.dll" /s)
:ends