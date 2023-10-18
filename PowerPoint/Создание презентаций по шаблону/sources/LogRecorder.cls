VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module      : LogRecorder
' Автор     : EducatedFool  (Игорь)                    Дата: 01.09.2010
' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
'---------------------------------------------------------------------------------------
Public Enum LOG_TYPE    ' типы лог-файлов
    LOG_CURRENT = 0
    LOG_DEFAULT = 1
    LOG_ENCRYPTED = 2
    LOG_SPECIAL = 4
    LOG_ERRORS = 8
    
    LOG_MAIL = 9
End Enum

Public Enum LOG_SEPARATOR_TYPE    ' добавление разделителей в лог
    LOG_SEPARATOR_NONE = 0
    LOG_SEPARATOR_BEFORE = 1
    LOG_SEPARATOR_AFTER = 2
    LOG_SEPARATOR_BEFORE_AND_AFTER = 3
End Enum

Public LogFileFullName As String
Public LogIndex As LOG_TYPE
Public index As Long

Public buffer As String
Public maxLogSizeBeforeSave As Long
Public LogSeparator As String


Private Sub Class_Terminate()
    buffer = "" 'CloseLog
End Sub

Private Sub Class_Initialize()
    LogSeparator = String(50, "=") & vbNewLine
    If maxLogSizeBeforeSave = 0 Then maxLogSizeBeforeSave = 5000
End Sub


Sub SaveLog()
    If AddIntoTXTfile(LogFileFullName, buffer) Then
        buffer = Empty
    Else
        Debug.Print "Ошибка записи лога в файл " & LogFileFullName
    End If
End Sub

Sub OpenLog()
    buffer = ""    ' очистка буфера
    SaveTXTfile LogFileFullName, ""    ' сохранение пустого файла
    AddRecord "НАЧИНАЕМ ЛОГИРОВАНИЕ...", , , LOG_SEPARATOR_BEFORE, True    ' стартовая запись
End Sub

Sub CloseLog()
    AddRecord "ЗАВЕРШАЕМ ЛОГИРОВАНИЕ...", , , LOG_SEPARATOR_AFTER, True    ' конечная запись
    buffer = ""    ' очистка буфера
End Sub

Sub Show()
    'On Error Resume Next
    Path$ = Chr(34) & LogFileFullName & Chr(34)
    CreateObject("wscript.shell").Run Path$
End Sub

Sub AddRecord(ByVal EventX As String, Optional ByVal Info As String, _
              Optional ByVal Level As Integer = 0, _
              Optional ByVal LogSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
              Optional ByVal ForceSavingLog As Boolean = False)

'    txt = String(451, " "): Mid(txt, 1) = Left(EventX, 400): Mid(txt, 51) = Info '28/02/2022
    txt = String(851, " "): Mid(txt, 1) = Left(EventX, 800): Mid(txt, 51) = Info
    'txt = String(Level, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(txt) & vbNewLine
    txt = Format(Now, "YYYY-MM-DD HH:NN:SS") & String(1 + Level, vbTab) & Trim(txt) & vbNewLine

    txt = IIf(LogSeparatorType = LOG_SEPARATOR_BEFORE Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "") & _
          txt & IIf(LogSeparatorType = LOG_SEPARATOR_AFTER Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "")

    buffer = buffer & txt
    If ForceSavingLog Or (Len(buffer) > maxLogSizeBeforeSave) Then SaveLog
End Sub


Sub SaveLogAs(ByVal filename As String, Optional ByVal Overwrite As Boolean = False)
    On Error Resume Next
    txt = ReadTXTfile(LogFileFullName)
    If Overwrite Then
        SaveTXTfile filename, txt
    Else
        AddIntoTXTfile filename, txt
    End If
End Sub

' ===============
Private Function ReadTXTfile(ByVal filename As String) As String
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(filename, 1, True): ReadTXTfile = ts.readall: ts.Close
    Set ts = Nothing: Set FSO = Nothing
End Function

Private Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set FSO = Nothing
End Function

Private Function AddIntoTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(filename, 8, True): ts.Write txt: ts.Close
    Set ts = Nothing: Set FSO = Nothing
    AddIntoTXTfile = Err = 0
End Function




