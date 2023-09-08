@echo off
setlocal ENABLEDELAYEDEXPANSION
	rem c:\exiv2\exiv2.exe -M"set Xmp.dc.subject "%substring%"" "%filename%"
set dirpath=%1
set taglist=%~2

for %%f in (%dirpath%) do (
	set filename=%%f
	call :parsetags
)	
	
:parsetags
SETLOCAL
:stringloop
	if "%taglist%" EQU "" goto END
	for /f "delims=;" %%a in ("%taglist%") do set substring=%%a
	rem echo %filename% %substring%
	c:\exiv2\exiv2.exe -M"set Xmp.dc.subject "%substring%"" "%filename%"
:striploop
    set stripchar=%taglist:~0,1%
    set taglist=%taglist:~1%
    if "%taglist%" EQU "" goto stringloop
    if "%stripchar%" NEQ ";" goto striploop
    goto stringloop
:END
ENDLOCAL