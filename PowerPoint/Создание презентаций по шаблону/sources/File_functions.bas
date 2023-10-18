Attribute VB_Name = "File_functions"
Private Declare PtrSafe Function StrCmpLogicalW Lib "shlwapi" _
(ByVal s1 As String, ByVal s2 As String) As Integer
'https://docs.microsoft.com/ru-ru/windows/win32/api/shlwapi/nf-shlwapi-strcmplogicalw?redirectedfrom=MSDN
'Compares two Unicode strings. Digits in the strings are considered as numerical content rather than text. This test is not case-sensitive.


'Работа с файловой системой
'https://vremya-ne-zhdet.ru/vba-excel/obyekt-filesystemobject/
'Добавил функцию сортировки списка файлов по тому же алгоритму, как это делает Windows
'Добавил диалго выбора файла
'26/01/2022 в GetAllSubFolders2 изменил порядок заполнения коллекции - теперь сперва вставляется путь корневой папки
Const sVersion As String = "Модуль работы файловой системой. Версия 4"

Function GetFolderPath(Optional ByVal Title As String = "Укажите директорию", _
                       Optional ByVal InitialPath As String = "c:\", _
                       Optional ByVal ButtonCaption As String = "Открыть") As String
    Dim PS As String: PS = "\"
    With ActivePresentation.Application.FileDialog(4) 'msoFileDialogFolderPicker
        If Not Right$(InitialPath, 1) = PS Then InitialPath = InitialPath & PS
        .ButtonName = ButtonCaption: .Title = Title: .InitialFileName = InitialPath
        If .Show <> -1 Then Exit Function
        GetFolderPath = .SelectedItems(1)
        If Not Right$(GetFolderPath, 1) = PS Then GetFolderPath = GetFolderPath & PS
    End With
End Function
 

Sub GetAllSubFolders2(ByVal sMainPath, ByRef cFolder As Collection)
Dim FSO As Object
Dim oMainFolder As Object
Dim oSubFolder As Object


Set FSO = CreateObject("Scripting.FileSystemObject")
Set oMainFolder = FSO.GetFolder(sMainPath)

For Each oSubFolder In oMainFolder.SubFolders
    cFolder.Add oSubFolder.Path & "\"
    GetAllSubFolders2 oSubFolder.Path, cFolder
Next oSubFolder

End Sub

Sub GetSortedFilesInFolder(ByVal sMainPath, ByRef aFileNames() As String, Optional ByVal sMask As String = "*")
Dim FSO As Object
Dim oMainFolder As Object
Dim oFile As Object
Dim nFilesCount As Integer
'Dim aFileNames() As String
'dim aFileArray as
Dim cFcount As Integer


Set FSO = CreateObject("Scripting.FileSystemObject")
Set oMainFolder = FSO.GetFolder(sMainPath)

nFilesCount = oMainFolder.Files.Count
ReDim aFileNames(1 To nFilesCount)

cFcount = 0

For Each oFile In oMainFolder.Files
    If oFile.Name Like "*" & sMask Then
        cFcount = cFcount + 1
        Namer$ = oFile.Name
        aFileNames(cFcount) = StrConv(Namer, vbUnicode)
    End If
Next

For AName = 1 To nFilesCount
    For BName = (AName + 1) To nFilesCount
        If StrCmpLogicalW(aFileNames(AName), aFileNames(BName)) = 1 Then
            buffer = aFileNames(BName)
            aFileNames(BName) = aFileNames(AName)
            aFileNames(AName) = buffer
        End If
    Next
Next
For i = 1 To nFilesCount
    aFileNames(i) = StrConv(aFileNames(i), vbFromUnicode)
    If i > 1 Then
    content = content & "," & aFileNames(i)
    Else
    content = aFileNames(i)
    End If
Next

End Sub

Function GetFilePath(Optional ByVal Title As String = "Укажите файл для обработки", _
                       Optional ByVal InitialPath As String = "c:\", _
                       Optional ByVal ButtonCaption As String = "Открыть") As String

With ActivePresentation.Application.FileDialog(3) 'msoFileDialogOpen
    .ButtonName = ButtonCaption: .Title = Title: .InitialFileName = InitialPath
    '.Filters.Add "Презентации", "*.ppt;*.pptx;*.pptm;*.pps;*.ppsx;*.odp"
    If .Show <> -1 Then Exit Function
    Let GetFilePath = .SelectedItems(1)
End With

End Function


Private Sub test()
Dim sResult As String
Dim lResult As Long

startTime = Timer
sResult = findFileInFolder("E:\Calibre", "metadata_db_prefs_backup.json")
lResult = Timer - startTime
MsgBox lResult & " - " & sResult

'MsgBox Dir("E:\Calibre\Calibre_portable\Calibre Portable\ebook-viewer-portable.exe")
End Sub


Function CreateFolder(folderPath As String)
Dim FSO As Object
Dim oMainFolder As Object
Dim oSubFolder As Object
Dim sFullPath As String
Set FSO = CreateObject("Scripting.FileSystemObject")

For Each sFolder In Split(folderPath, "\")
    sFullPath = sFullPath & sFolder & "\"
    If Not FSO.FolderExists(sFullPath) Then
        FSO.CreateFolder (sFullPath)
    End If
Next

If FSO.FolderExists(folderPath) Then
    CreateFolder = folderPath
Else
    CreateFolder = ""
End If

'Set oMainFolder = FSO.GetFolder(sMainPath)
End Function



Function findFileInFolder(sMainFolder As String, sFileName As String)

Dim StrFile As String
Dim cFolderList As New Collection
'Dim sSubFolder As String

GetAllSubFolders2 sMainFolder, cFolderList

For Each sSubFolder In cFolderList
    StrFile = Dir(sSubFolder & sFileName)
    If StrFile = sFileName Then
        findFileInFolder = sSubFolder & sFileName
        Exit Function
    End If
Next sSubFolder

End Function


Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    
    Set oExec = oShell.exec(sCmd)
    Set oOutput = oExec.stdout

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.readline
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function

Public Function bFileExists2(sFile As String) As Boolean
    'Проверка наличия файла.
    'Если не работает, подключить References -> Microsoft Scripting Runtime
    Dim f As New Scripting.FileSystemObject
    If f.fileexists(sFile) = True Then
        bFileExists2 = True
    Else
        bFileExists2 = False
    End If

End Function

Public Function bFileExists(sFile As String) As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.fileexists(sFile) Then
        bFileExists = True
    Else
        bFileExists = False
    End If

End Function

Public Function DeleteFilesByMask(sFolder As String, Optional sMaks As String = "*.*")


End Function

Private Sub test2()
MsgBox ShellRun("cmd.exe /c dir")
End Sub
