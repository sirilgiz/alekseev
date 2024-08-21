# Версия 1.0.20240819
# Скрипт копирует файлы книг из указанной папки с переименованием на основе данных из файла со списком книг
# Скрипт ожидает получить файл, содержащий данные в формате csv с разделителем ';' и кодировкой UTF-8 со следующими полями: 
#   Название;Авторы;Путь к файлу. 
#   Первая строка файла содержит заголовки полей, вторая и последующая - записи о книгах. 
#   Одна запись о книге может содержать 1 и более путей к файлам, разделенные символом '|'

$Logfile = "bookrenamer.log"
function WriteLog
{
Param ([string]$LogString)

$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage
}

WriteLog "=================НАЧАЛО ОБРАБОТКИ================="

# Диалог выбора папки
Add-Type -AssemblyName System.Windows.Forms
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "УКАЖИТЕ ПАПКУ С ФАЙЛАМИ КНИГАМИ"
$folderBrowser.SelectedPath = (Get-Location).Path
$null = $folderBrowser.ShowDialog()
$bookfilefolder = $folderBrowser.SelectedPath


Write-Output "ПАПКА С КНИГАМИ: $bookfilefolder"
WriteLog "ПАПКА С КНИГАМИ: $bookfilefolder"

#$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "УКАЖИТЕ ПАПКУ ДЛЯ ПЕРЕИМЕНОВЫХ КНИГ"
$folderBrowser.SelectedPath = (Get-Location).Path
$null = $folderBrowser.ShowDialog()
$renamedbookfilefolder = $folderBrowser.SelectedPath

Write-Output "ПАПКА ДЛЯ ПЕРЕМЕНОВАННЫХ КНИГ: $renamedbookfilefolder"
WriteLog "ПАПКА ДЛЯ ПЕРЕМЕНОВАННЫХ КНИГ: $renamedbookfilefolder"

#Диалог выбора файла

#Ожидаемый формат записи: Название;Автор;Путь к файлу

$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$fileBrowser.Filter = "csv СПИСОК КНИГ (*.csv) | *.csv"
$fileBrowser.InitialDirectory = (Get-Location).Path
$null = $fileBrowser.ShowDialog()
$csvfilefullpath = $fileBrowser.FileName
Write-Output "ФАЙЛ СО СПИСОКОМ КНИГ: $csvfilefullpath"
WriteLog "ФАЙЛ СО СПИСОКОМ КНИГ: $csvfilefullpath"


$csv = (get-content $csvfilefullpath -Encoding UTF8)
$csvlen = $csv.Count
$b = 0

WriteLog "НАЧАЛИ ОБРАБОТКУ ФАЙЛА СО СПИСКОМ КНИГ"

foreach ($file in $csv) {
    $b = $b + 1
    if ($b -eq 1) #пропускаем первую строку с заголовками
    {
        WriteLog "ЗАГОЛОВКИ: $file"    
        continue 
    }
    WriteLog "  запись $b из $csvlen '$file'"
    Write-Progress -Activity "ОБРАБОТКА КНИГ" -Status "обработка записи $b из $csvlen" -Id 1 -PercentComplete ([int][Math]::Round(($b/$csvlen)*100,0))
    $line = $file.Split(";")
    if ($line.Length -lt 3) {
        WriteLog "    oшибочная запись, пропускаем"
        Write-Output "Wrong lenght: "+$line+"->"+$file
        continue
    }
    $bookname=$line[0]
    $bookauthor=$line[1]
    $files=$line[2]

    if ($bookauthor.Length -gt 0) {
        $bookauthor = " - "+$bookauthor.Trim()
    }
    else
    {
        $bookauthor=" - Нет автора"
    }
    
    
    $filenames = $files.Split("|")
    foreach ($filename in $filenames)
    {
        $oldfilename = $filename.Split("\")[-1]
        $extpos = $oldfilename.LastIndexOf(".")
        if ($extpos -gt -1)
        {
            $ext = $oldfilename.Split(".")[-1]
            $oldfilename = $oldfilename.Substring(0,$extpos) 
            #$oldfilenamelen = $oldfilename.Length
            #$ext=$oldfilename.Substring($extpos,$oldfilenamelen-$extpos) 
        }
        else
        {
            $ext = ""
        }
        
        
        $newfilename=$bookname+$bookauthor+"."+$ext
        WriteLog "    удаляем запрещенные символы из имени файла: $newfilename"
        $newfilename = $newfilename -replace "[^\w \.-]", ""
        $newfilename = $newfilename.Trim()
        $newfilefullpath = $renamedbookfilefolder+"\"+$newfilename
        WriteLog "      новое имя файла: $newfilename"

        if ([System.IO.File]::Exists($newfilefullpath))
        {
            WriteLog "    такой файл уже существует. Переименуем: $newfilename"
            $newfilename=$bookname+$bookauthor+" ("+$oldfilename+")."+$ext
            $newfilename = $newfilename -replace "[—–]", "-"
            $newfilename = $newfilename -replace "[^\w \.,-»«\)\(]", ""
            $newfilename = $newfilename.Trim()
            WriteLog "      новое имя файла: $newfilename"
            $newfilefullpath = $renamedbookfilefolder+"\"+$newfilename
        }

        $oldfilefullpath=$bookfilefolder+"\"+$oldfilename+"."+$ext
        #$cmd = $bookfilefolder+"\"+$oldfilename+"."+$ext => "+$newfilefullpath
        WriteLog "  копируем $oldfilename.$ext в $newfilefullpath"
        Copy-Item $oldfilefullpath -Destination $newfilefullpath
       
        
    }
}
WriteLog "=================КОНЕЦ ОБРАБОТКИ=================="    