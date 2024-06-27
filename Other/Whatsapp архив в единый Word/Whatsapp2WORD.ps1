$attachmentRegex = "<прикреплено:\s*(.*?)>"
$datestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$documentName = "WhatsApp2Word_$Datestamp.doc"


# Диалог выбора папки
Add-Type -AssemblyName System.Windows.Forms
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Выберите общую папку, в которой сохранены подпапки с архивами чатов"
$folderBrowser.SelectedPath = (Get-Location).Path
$null = $folderBrowser.ShowDialog()
$selectedFolder = $folderBrowser.SelectedPath

#Создаем документ
$Word = New-Object -ComObject Word.Application
$Word.Visible = $false
$Document = $Word.Documents.Add()

$DocumentPath = $selectedFolder + $documentName

# Пробегаемся по всем папкам в поисках чатов
$chatfiles = Get-ChildItem -Path $selectedFolder -Recurse -File -Filter "_chat.txt"
$c = 0
$totalchats = $chatfiles.Count

foreach ($chat in $chatfiles) {
    $c = $c + 1
    Write-Progress -Activity "ОБРАБОТКА ЧАТОВ" -Status "обрабатывается $c из $totalchats" -Id 1 -PercentComplete ([int][Math]::Round(($c/$totalchats)*100,0))
    $Document.SaveAs($DocumentPath)
    $Paragraph = $Document.Paragraphs.Add()
    $Paragraph = $Document.Paragraphs($Document.Paragraphs.Count())
    $Paragraph.Range.Text = $chat.Directory.Name
    
    $Paragraph.Range.Style = "Заголовок 1"

    $Paragraph.Range.Font.Size = 18
    
    $chattext = Get-Content $chat.FullName -Encoding UTF8
    $totalmessage = $chattext.Count
    $i = 0
    foreach ($chatmessage in $chattext)  {
        $i = $i + 1
        Write-Progress -Activity "ОБРАБОТКА СООБЩЕНИЙ В ЧАТЕ $($chat.Directory.Name)" -Status "обработано $i из $totalmessage" -Id 2 -PercentComplete ([int][Math]::Round(($i/$totalmessage)*100,0)) -ParentId 1
        $Paragraph = $Document.Paragraphs.Add()
        $Paragraph = $Document.Paragraphs($Document.Paragraphs.Count())
        $Paragraph.Range.Font.Size = 12
        $Paragraph.Range.Style = "Обычный"
        $Paragraph.Range.Text = $chatmessage

        if ($chatmessage -match $attachmentRegex) {
            $attachmentName = $matches[1]
            $attachmentRelPath = "."+ $selectedFolder.Substring($selectedFolder.LastIndexOf("\")) +"\"+ $chat.Directory.Name + "\" + $attachmentName
            $attachmentFullPath = $chat.Directory.FullName + "\" + $attachmentName
            $Paragraph = $Document.Paragraphs.Add()
            $Paragraph = $Document.Paragraphs($Document.Paragraphs.Count())
            $Paragraph.Range.Text = $attachmentName
            $Selection = $Paragraph.Range
            
            
            if ($attachmentName.EndsWith(".jpg") -or $attachmentName.EndsWith(".jpeg") -or $attachmentName.EndsWith(".png")) {
                $hyperlink = $Document.Hyperlinks.Add($Paragraph.Range,$attachmentRelPath)                 
                $picture = $Selection.InlineShapes.AddPicture($attachmentFullPath) #, $false, $false, $Paragraph.Range)
                $picture.LockAspectRatio = $true
                $picture.Height = 100
            }
            else {
                $hyperlink = $Document.Hyperlinks.Add($Paragraph.Range, $attachmentRelPath)
            }

        }
        else {
            $Paragraph.Range.Text = $chatmessage
        }

    }

}
#Добавляем оглавление
$null = $Document.TablesOfContents.Add($Document.Range(0,0))

$Document.SaveAs($DocumentPath)
$Word.Visible = $true
#$Document.Close()
#$Word.Quit()
