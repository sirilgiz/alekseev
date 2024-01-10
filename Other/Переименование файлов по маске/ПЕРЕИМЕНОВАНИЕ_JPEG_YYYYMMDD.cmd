@echo off
echo. 
echo. 
echo  ВНИМАНИЕ!                                       
echo   ДАННЫЙ СКРИПТ ОБРАБАТЫВАЕТ ВСЕ ФАЙЛЫ, ПО МАСКЕ "минимум 15 СИМВОЛОВ + . + РАСШИРЕНИЕ", В ТЕКУЩЕЙ И ВСЕХ ВЛОЖЕННЫХ ПАПКАХ И ПЕРЕИМЕНОВЫВАЕТ ИХ ПО ПРАВИЛУ:
echo     XXXXXXXXXXXXXXX.YYY	===	XXXX-XX-XX XXXXXX X.YYY
echo     XXXXXXXXXXXXXXX.YYYY	===	XXXX-XX-XX XXXXXX X.YYYY
echo.
echo  ВАЖНО!                                      
echo   ЕСЛИ ФАЙЛ С НОВЫМ НАЗВАНИЕМ УЖЕ СУЩЕСТВУЕТ, СКРИПТ ЕГО ПЕРЕЗАПИШЕТ!                     
echo.                                             
echo.
@set /p answer=ДЛЯ ПРОДОЛЖЕНИЯ ВВЕДИТЕ 1 + ENTER: 
if %answer%==1 (
forfiles /S /M ????????????*.jp*g /C "cmd /v:on /c set f=@file && set f=!f: =! && set f=!f:-=! && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9,4! !f:~12!' && set n=!n:~1,-2! && move /Y @path \"!n!"
forfiles /S /M ???????????.jp*g /C "cmd /v:on /c set f=@file && set f=!f: =! && set f=!f:-=! && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9!' && set n=!n:~1,-2! && move /Y @path \"!n!"
forfiles /S /M ??????????.jp*g /C "cmd /v:on /c set f=@file && set f=!f: =! && set f=!f:-=! && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9!' && set n=!n:~1,-2! && move /Y @path \"!n!"
forfiles /S /M ?????????.jp*g /C "cmd /v:on /c set f=@file && set f=!f: =! && set f=!f:-=! && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9!' && set n=!n:~1,-2! && move /Y @path \"!n!"
forfiles /S /M ????????.jp*g /C "cmd /v:on /c set f=@file && set f=!f: =! && set f=!f:-=! && set n='!f:~1,4!-!f:~5,2!-!f:~7!' && set n=!n:~1,-2! && move /Y @path \"!n!"
)
pause
rem if %answer%==1 forfiles /S /M ???????????????.jp*g /C "cmd /v:on /c set f=@file && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9!' && set n=!n:~1,-2! && move /Y @path \"!n!"
rem forfiles /S /M ???????????????.jp*g /C "cmd /v:on /c set f=@file && set n='!f:~1,4!-!f:~5,2!-!f:~7,2! !f:~9!' && set n=!n:~1,-2! && rename @path \"!n!"
rem длина имени файла минимум 15 символов, но может быть и больше через маску ???????????????*.jp*g
rem удаляю пробелы через set n=!n: =! и дефисы через set n=!n:-=! -- это позволит безболезненно обработать уже обработанные файлы (так как удалит все пробелы  и дефисы и заново их добавит