Attribute VB_Name = "Text_functions"
Public Function SHA1(ByVal s As String) As String
    Dim Enc As Object, Prov As Object
    Dim Hash() As Byte, i As Integer

    Set Enc = CreateObject("System.Text.UTF8Encoding")
    Set Prov = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Set Prov = CreateObject("System.Security.Cryptography.SHA256Managed")
    

    Hash = Prov.ComputeHash_2(Enc.GetBytes_4(s))

    SHA1 = ""
    For i = LBound(Hash) To UBound(Hash)
        SHA1 = SHA1 & Hex(Hash(i) \ 16) & Hex(Hash(i) Mod 16)
    Next
End Function


Function GetHash(ByVal txt$) As String
    Dim oUTF8, oMD5, abyt, i&, k&, hi&, lo&, chHi$, chLo$
    Set oUTF8 = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    abyt = oMD5.ComputeHash_2(oUTF8.GetBytes_4(txt$))
    For i = 1 To LenB(abyt)
        k = AscB(MidB(abyt, i, 1))
        lo = k Mod 16: hi = (k - lo) / 16
        If hi > 9 Then chHi = Chr(Asc("a") + hi - 10) Else chHi = Chr(Asc("0") + hi)
        If lo > 9 Then chLo = Chr(Asc("a") + lo - 10) Else chLo = Chr(Asc("0") + lo)
        GetHash = GetHash & chHi & chLo
    Next
    Set oUTF8 = Nothing: Set oMD5 = Nothing
End Function


Function SplitText(sSentense As String) As String
'Функция принимает строку (предложение), делит на слова, разделенные точкой с запятой ";"
Dim aClearString() As String
Dim sTemp As String
Dim FReg As Object

'приводим строку к нижнему регистру
sTemp = StrConv(sSentense, vbLowerCase)

If FReg Is Nothing Then
  Set FReg = CreateObject("VBScript.RegExp")
  With FReg
    .MultiLine = False
    .Global = True
    .IgnoreCase = True
    .Pattern = "[^ da-zёа-я0-9]"
  End With
  
sTemp = Trim(FReg.Replace(sTemp, ""))

Do While InStr(sTemp, "  ")
  sTemp = Replace(sTemp, "  ", " ")
Loop

End If
'удаляем знаки препинания и прочую аттрибутику
'stop_symbols = ".,!?:;-()<>/\\\\'1234567890[]=+\#%&"

aClearString = Split(sTemp, " ")
'aClearString = Filter(aClearString, "")
SplitText = Join(aClearString, ";") & ";"
'про массивы и словари
'http://perfect-excel.ru/publ/excel/makrosy_i_programmy_vba/ischerpyvajushhee_opisanie_obekta_dictionary/7-1-0-101


End Function

Function SplitText2(sSentense As String) As String
'Функция принимает строку (предложение), делит на слова, разделенные точкой с запятой ";"
Dim aClearString() As String
Dim sTemp As String
Dim FReg As Object

'приводим строку к нижнему регистру
sTemp = StrConv(sSentense, vbLowerCase)

If FReg Is Nothing Then
  Set FReg = CreateObject("VBScript.RegExp")
  With FReg
    .MultiLine = False
    .Global = True
    .IgnoreCase = True
    .Pattern = "[^ da-zёа-я0-9\\]"
  End With
  
sTemp = Trim(FReg.Replace(sTemp, ""))

Do While InStr(sTemp, "  ")
  sTemp = Replace(sTemp, "  ", " ")
Loop

sTemp = Replace(sTemp, " ", "_")

End If
'удаляем знаки препинания и прочую аттрибутику
'stop_symbols = ".,!?:;-()<>/\\\\'1234567890[]=+\#%&"

aClearString = Split(sTemp, "\")
'aClearString = Filter(aClearString, "")
SplitText2 = Join(aClearString, ";")
'про массивы и словари
'http://perfect-excel.ru/publ/excel/makrosy_i_programmy_vba/ischerpyvajushhee_opisanie_obekta_dictionary/7-1-0-101


End Function

Function EncodeUTF8(s)
    Dim i, c As Long, utfc, b1, b2, b3
    
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
 
        If c < 128 Then
            utfc = Chr(c)
        ElseIf c < 2048 Then
            b1 = c Mod &H40
            b2 = (c - b1) / &H40
            utfc = Chr(&HC0 + b2) & Chr(&H80 + b1)
        ElseIf c < 65536 And (c < 55296 Or c > 57343) Then
            b1 = c Mod &H40
            b2 = ((c - b1) / &H40) Mod &H40
            b3 = (c - b1 - (&H40 * b2)) / &H1000
            utfc = Chr(&HE0 + b3) & Chr(&H80 + b2) & Chr(&H80 + b1)
        Else
            ' ??????? ??? ??????? ???????? UTF-16
            utfc = Chr(&HEF) & Chr(&HBF) & Chr(&HBD)
        End If

        EncodeUTF8 = EncodeUTF8 + utfc
    Next
End Function


Private Sub test_split()
MsgBox SHA1("Профессиональная деформация. Профессиональная деформация юристов и государственных служащих.pptx")
End Sub
