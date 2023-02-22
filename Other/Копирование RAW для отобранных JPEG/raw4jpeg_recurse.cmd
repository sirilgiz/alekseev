@ECHO off
SETLOCAL enableextensions 
SETLOCAL enabledelayedexpansion

:: 1.0 от 20/02/2023 — первая версия
:: 1.1 от 22/02/2023 — обновил алгоритм поиска для работы с вложенными папками

ECHO Скрипт поиска ARW-файлов для одноимённых JPEG-файлов. 
ECHO.
ECHO ВАЖНО: файлы ARW и JPEG должны иметь одинаковые имена, т.е. для файла PHOTO_123.ARW будет искаться файл PHOTO_123.JPEG
ECHO ВАЖНО: найденные ARW-файлы будут ПЕРЕМЕЩЕНЫ в отдельную папку, т.е. БУДУТ УДАЛЕНЫ из исходной папки (это практически мгновенно независимо от размера файла)
ECHO.
ECHO [Рабочая_папка]................Отдельная папка, в которой будет работать скрипт (название неважно)
ECHO    └AARW.......................Сюда нужно скопировать все папки с файлами в формате ARW (создаётся автоматически)
ECHO    └JJPEG......................Сюда нужно скопировать все папки с файлами в формате JPEG (создаётся автоматически)
ECHO    └RAW........................Сюда перенесутся все найденные файлы (создаётся автоматически)
ECHO    └raw4jpeg_recurse.cmd.......Скрипт     
ECHO.
ECHO ВАЖНО: структуры каталогов AARW и JJPEG должны совпадать, т.к. скрипт ищет по полному соответствию пути и имени файла (разница только в расширении)
PAUSE

SET workingPath=%~dp0test1\
SET currentPath=%workingPath%
SET currentARWFile="none"
SET currentJPEGFile="none"
SET currentRAWFile="none"

ECHO ----
ECHO НАЧИНАЕМ ОБРАБОТКУ 
ECHO ---- >> log.txt
ECHO %date% %time% НАЧИНАЕМ ОБРАБОТКУ >> log.txt

ECHO Рабочая папка: %workingPath%
ECHO Рабочая папка: %workingPath% >> log.txt

:: Если папок JJPEG и AARW нет, создаём их и выходим (так как пользователь должен поместить туда файлы)
IF NOT EXIST "%workingPath%JJPEG\" (
    MKDIR "%workingPath%JJPEG"
    ECHO Папка JJPEG не найдена и создана автоматически. Скопируйте в эту папку отобранные JPEG-файлы
    IF NOT EXIST "%workingPath%AARW\" (
        MKDIR "%workingPath%AARW"
        ECHO Папка AARW не найдена и создана автоматически. Скопируйте в эту папку исходные ARW-файлы
    )
    goto exit
)

for /r "%workingPath%JJPEG\" %%n in ("*.JP*G") do (
    :: Если обрабатываем новую директорию, напишем в лог/выведем на экран путь к ней
    IF "%%~dn%%~pn" NEQ "!currentPath!" (
        SET currentPath="%%~dn%%~pn"
        ECHO.
        ECHO Обрабатываем папку: !currentPath!
        ECHO. >>log.txt
        ECHO Обрабатываем папку: %%~pn >>log.txt
    )
    :: ECHO файл --- %%~dn%%~pn%%~nn
    SET currentJPEGFile=%%~dn%%~pn%%~nn
    ::Проверяем, что файл ARW для данного JPEG существует
    ECHO    └Обрабатываем JPEG: %%~nn%%~xn
    ECHO    └Обрабатываем JPEG: %%~nn%%~xn >> log.txt
    
    SET currentARWFile=!currentJPEGFile:JJPEG=AARW!.ARW
    SET currentRAWFile=!currentJPEGFile:JJPEG=RAW!.ARW
    SET currentRAWPath=!currentPath:JJPEG=RAW!

    IF EXIST "!currentARWFile!" (
        :: Если файл ARW найден, нужно создать папку RAW (если она не существует) 
		IF NOT EXIST "!currentRAWPath!" (
            MKDIR ""!currentRAWPath!""
        )
		:: Переносим ARW-файл в папку RAW
        ECHO        └ARW найден, переносим
        ECHO        └ARW найден, переносим >> log.txt
        MOVE /Y "!currentARWFile!" "!currentRAWFile!" >> log.txt
    ) ELSE (
        :: иначе если ARW не найден
        ECHO        └ARW не найден: "!currentARWFile!"
        ECHO        └ARW не найден: "!currentARWFile!" >> log.txt
    )
    
)

:exit
ECHO ОБРАБОТКА ЗАВЕРШЕНА
ECHO %date% %time% ОБРАБОТКА ЗАВЕРШЕНА >> log.txt
ECHO ---- >> log.txt