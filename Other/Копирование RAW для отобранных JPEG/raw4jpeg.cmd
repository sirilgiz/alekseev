@echo off
goto start
------
��ਯ� ��ॡ�ࠥ� 䠩�� � ⥪�饩 ����� � ��� ������� �����㦥����� JPEG-䠩�� ��� ���������� RAW-䠩� � ����� �� 1 �஢��� ���. �᫨ ᮮ⢥�����騩 RAW-䠩� ������, �� ��७����� � ����� RAW (ᮧ������ ��⮬���᪨).
�� ����, �᫨ � ��� ���� ����� Photo, � ���ன ��⥭樠�쭮 ��室���� RAW-䠩�� � ���� ����� JPEG � �⮡࠭�묨 ���ࠬ�, �㦭� ᪮��஢��� ����� �ਯ� � ����� JPEG, � ᠬ� ����� JPEG �������� � ����� Photo, ��᫥ 祣� �������� �ਯ�.

Photo               << ����� �㤥� �᪠�� RAW-䠩�� (�������� ����� �� ����� ���祭��)
|_RAW               << � ��७������ �������� RAW-䠩�� (�� ����� ᮧ������ ��⮬���᪨)
|_JPEG              << ����� �⮡࠭�� JPEG-䠩�� (�������� ����� �� ����� ���祭��)
  |_raw4jpeg.cmd    << �� ��� �ਯ� (�������� �ਯ� �� ����� ���砭��)
------
v. 1.0 2022-12-07 
------

:start
:: ��ࠬ����, ����� ����� ������:
SET raw_file_ext=ARW& :: ���७�� RAW-䠩���. ���ਬ��, ��ப� "SET raw_file_ext=ARW" ������ ���७�� .ARW
SET raw_new_folder_name=RAW& :: ���।���� �������� �����, � ������ �ਯ� ��७��� �������� RAW-䠩��. ���ਬ��, ��ப� "SET raw_new_folder_name=RAW" ������ �������� ����� "RAW"

for %%f in (*.jp*g) do (
	IF EXIST ..\%%~nf.%raw_file_ext% (
		IF NOT EXIST ..\%raw_new_folder_name% md ..\%raw_new_folder_name%
		move /Y ..\%%~nf.%raw_file_ext% ..\%raw_new_folder_name%\%%~nf.%raw_file_ext%
        ) ELSE (
            echo %%~nf.%raw_file_ext% �� ������
        )
)

pause