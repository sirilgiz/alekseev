@echo off
echo. 
echo. 
echo  ��������!                                       
echo   ������ ������ ������������ ��� �����, �� ����� "������ 15 �������� + . + ����������", � ������� � ���� ��������� ������ � ��������������� �� �� �������:
echo     XXXXXXXXXXXXXXX.YYY	===	XXXX-XX-XX XXXXXX X.YYY
echo     XXXXXXXXXXXXXXX.YYYY	===	XXXX-XX-XX XXXXXX X.YYYY
echo.
echo  �����!                                      
echo   ���� ���� � ����� ��������� ��� ����������, ������ ��� �����������!                     
echo.                                             
echo.
@set /p answer=��� ����������� ������� 1 + ENTER: 
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
rem ����� ����� 䠩�� ������ 15 ᨬ�����, �� ����� ���� � ����� �१ ���� ???????????????*.jp*g
rem 㤠��� �஡��� �१ set n=!n: =! � ����� �१ set n=!n:-=! -- �� �������� ������������� ��ࠡ���� 㦥 ��ࠡ�⠭�� 䠩�� (⠪ ��� 㤠��� �� �஡���  � ����� � ������ �� �������