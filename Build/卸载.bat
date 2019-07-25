@echo off
if %PROCESSOR_ARCHITECTURE% == AMD64 (
if not exist "%windir%\Syswow64\MiniblinkSDK.dll" goto ends
regsvr32 "%windir%\Syswow64\MiniblinkSDK.dll" /s /u
del /f /s /q "%windir%\Syswow64\MiniblinkSDK.dll"
) else (
if not exist "%windir%\System32\MiniblinkSDK.dll" goto ends
regsvr32 "%windir%\System32\MiniblinkSDK.dll" /s /u
del /f /s /q "%windir%\System32\MiniblinkSDK.dll"
)
:ends