VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   13575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18375
   OleObjectBlob   =   "UserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    #If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    #Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    #End If


Function ShowSlidePreview(CurrentSlide As Integer)

If (Not Not SlideArray) = 0 Then
       'массив пуст и записей в нем не было
Else
        If UBound(SlideArray) = 1 Then
            imgSlide1.Picture = LoadPicture(SlideArray(1))
            imgSlide2.Picture = Nothing
            imgSlide3.Picture = Nothing
        ElseIf UBound(SlideArray) = 2 Then
            imgSlide1.Picture = LoadPicture(SlideArray(1))
            imgSlide2.Picture = LoadPicture(SlideArray(2))
            imgSlide3.Picture = Nothing
        ElseIf UBound(SlideArray) = 3 Then
            imgSlide1.Picture = LoadPicture(SlideArray(CurrentSlide - 2))
            imgSlide2.Picture = LoadPicture(SlideArray(CurrentSlide - 1))
            imgSlide3.Picture = LoadPicture(SlideArray(CurrentSlide))
        Else
            If CurrentSlide >= 3 Then
                imgSlide1.Picture = LoadPicture(SlideArray(CurrentSlide - 2))
                imgSlide2.Picture = LoadPicture(SlideArray(CurrentSlide - 1))
                imgSlide3.Picture = LoadPicture(SlideArray(CurrentSlide))
            Else
                imgSlide1.Picture = LoadPicture(SlideArray(1))
                imgSlide2.Picture = LoadPicture(SlideArray(2))
                imgSlide3.Picture = LoadPicture(SlideArray(3))
            End If
        End If
End If


End Function


Private Function UpdateForm(Optional ByVal frmSlideText1Visible As Boolean = True, Optional ByVal frmSlideText2Visible As Boolean = True, Optional ByVal frmSlideImageVisible As Boolean = True)

If frmSlideText1Visible Then
    frmSlideText1.Visible = True
Else
    frmSlideText1.Visible = False
    txtSlideText1.Text = ""
End If

If frmSlideText2Visible Then
    frmSlideText2.Visible = True
Else
    frmSlideText2.Visible = False
    txtSlideText2.Text = ""
End If

If frmSlideImageVisible Then
    frmSlideImage.Visible = True
Else
    frmSlideImage.Visible = False
    imgSlideImage.Picture = Nothing
    lstFilesList.ListIndex = -1
End If

End Function

Private Function SavePresentationAs() As Boolean
Dim fDialog As FileDialog
Dim sImgFullPath As String

Set fDialog = Application.FileDialog(msoFileDialogSaveAs)

With fDialog
    .InitialFileName = MacrosPresentation.Path & "\" & TempPresentationName
    '.Filters.Add "Презентация PPTX", "*.pptx"
End With
    
If fDialog.Show = -1 Then
    sImgFullPath = fDialog.SelectedItems(1)
    TempPresentation.SaveCopyAs sImgFullPath
    TempPresentation.Save
    SavePresentationAs = True
End If

End Function
Private Sub btnSave_Click()
SavePresentationAs
End Sub


Private Sub frmPreview_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

End Sub

Private Sub Label1_Click()
Dim oShell As Object
Dim sPath As String
Let sPath = Environ("TMP")
Set oShell = CreateObject("Shell.Application")
oShell.Explore (sPath)


End Sub

Private Sub ScrollBar1_Change()
    ShowSlidePreview ScrollBar1.Value
End Sub

Private Sub UserForm_Initialize()
UserForm.Caption = sVersion
Set MacrosPresentation = ActivePresentation
SlideData.Template = 5

End Sub


Private Sub btnInsertSlide_Click()
'UpdateForm
ResetSlideData
CreateSlide
btnSave.Enabled = True
ScrollBar1.Max = UBound(SlideArray)
ScrollBar1.Value = UBound(SlideArray)
If frmSlideText1.Visible Then If chkClearText1.Value = True Then txtSlideText1.Text = ""
If frmSlideText2.Visible Then If chkClearText2.Value = True Then txtSlideText2.Text = ""
If frmSlideImage.Visible Then
    If chkHideUsedImg.Value = True Then
        If lstFilesList.ListIndex >= 0 Then
            lstFilesList.RemoveItem lstFilesList.ListIndex
            lstFilesList.ListIndex = -1
            imgSlideImage.Picture = Nothing
        End If
            
    End If
End If
End Sub



Private Sub lstFilesList_Click()
    SlideData.ImgFullPath = lstFilesList.Value
    Set imgSlideImage.Picture = ThumbnailPicFromFile(SlideData.ImgFullPath, 256, 256)
End Sub

Private Sub btnChooseImageFile_Click()
Dim fDialog As FileDialog
Dim sImgFullPath As String

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

With fDialog
    .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp; *.png", 1
    .ButtonName = "Выбрать"
    .Title = "Выберите изображение для вставки на слайд. " & sVersion
End With

If fDialog.Show = -1 Then
    sImgFullPath = fDialog.SelectedItems(1)
'sImgFullPath = GetFilePath("Укажите изображение для вставки", MacrosPresentation.Path, "Вставить")
    lstFilesList.AddItem sImgFullPath
    SlideData.ImgFullPath = sImgFullPath
    lstFilesList.Selected(lstFilesList.ListCount - 1) = True
    'Set imgSlideImage.Picture = ThumbnailPicFromFile(SlideData.ImgFullPath, 256, 256)
End If




End Sub

Private Sub btnChooseImageFolder_Click()
Dim FSO As Object, oSubFolder As Object, oFile As Object
Dim oPic As StdPicture
Dim imgFormats As Variant
Let imgFormats = Array("gif", "jpg", "jpeg", "bmp", "png")
Dim imgFormat As String
Dim i As Integer


SlideData.ImgFolderPath = GetFolderPath("Укажите папку с изображениями", MacrosPresentation.Path, "Выбрать")
Set FSO = CreateObject("Scripting.FileSystemObject")


If FSO.FolderExists(SlideData.ImgFolderPath) Then
    lstFilesList.Clear
    
    For Each oFile In FSO.GetFolder(SlideData.ImgFolderPath).Files
        For i = LBound(imgFormats) To UBound(imgFormats)
            imgFormat = FSO.GetExtensionName(oFile)
            If InStr(1, UCase(imgFormat), UCase(imgFormats(i)), vbTextCompare) > 0 Then
                lstFilesList.AddItem FSO.GetAbsolutePathName(oFile)
                'lstFilesList.AddItem FSO.GetFileName(oFile)
                Exit For
            End If
        Next i
    Next oFile
    If lstFilesList.ListCount Then
        lstFilesList.Selected(0) = True
        SlideData.ImgFullPath = lstFilesList.List(0)
        'Set imgSlideImage.Picture = ThumbnailPicFromFile(SlideData.ImgFolderPath & SlideData.ImgFileName, 256, 256)
        Set imgSlideImage.Picture = ThumbnailPicFromFile(SlideData.ImgFullPath, 256, 256)
        
        'SlideData.ImgFullPath = SlideData.ImgFolderPath & lstFilesList.List(0)
        'SlideData.ImgFullPath = lstFilesList.List(0)
        'SlideData.imgH = imgSlideImage.Picture.Height
        'SlideData.imgW = imgSlideImage.Picture.Width
    End If
End If



End Sub

Private Sub btnClearTextBox_Click()
txtSlideText.Text = ""
End Sub

Private Sub imgSlideImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
btnChooseImageFile_Click
End Sub
Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate1.Value = 0 Then optTemplate1.Value = 1
End Sub
Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate2.Value = 0 Then optTemplate2.Value = 1
End Sub
Private Sub Image3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate3.Value = 0 Then optTemplate3.Value = 1
End Sub
Private Sub Image4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate4.Value = 0 Then optTemplate4.Value = 1
End Sub
Private Sub Image5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate5.Value = 0 Then optTemplate5.Value = 1
End Sub
Private Sub Image6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate6.Value = 0 Then optTemplate6.Value = 1
End Sub
Private Sub Image7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate7.Value = 0 Then optTemplate7.Value = 1
End Sub
Private Sub Image8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If optTemplate8.Value = 0 Then optTemplate8.Value = 1
End Sub




Private Sub optTemplate1_Click()
UpdateForm True, False, False
End Sub

Private Sub optTemplate2_Click()
UpdateForm True, False, False
End Sub

Private Sub optTemplate3_Click()
UpdateForm True, False, False
End Sub

Private Sub optTemplate4_Click()
UpdateForm True, False, False
End Sub

Private Sub optTemplate5_Click()
UpdateForm True, False, False
End Sub

Private Sub optTemplate6_Click()
UpdateForm True, False, True
End Sub

Private Sub optTemplate7_Click()
UpdateForm True, True, False
End Sub

Private Sub optTemplate8_Click()
UpdateForm True, False, True
End Sub



Private Sub txtSlideText1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
txtSlideText1.Text = UCase(txtSlideText1.Text)
End Sub
Private Sub txtSlideText2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
txtSlideText2.Text = UCase(txtSlideText2.Text)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim msgUserAnswer As Variant
Dim sPresentationFullPath As String

If Not TempPresentation Is Nothing Then
    If Not TempPresentation.Saved Then
        msgUserAnswer = MsgBox("Презентация не сохранена! Сохранить?", vbQuestion + vbYesNoCancel, sVersion)
        Select Case msgUserAnswer
        Case vbNo
            TempPresentation.Close
            Kill TempFolderPath & TempPresentationName & "*.*"
        Case vbYes
            If SavePresentationAs = True Then
                TempPresentation.Close
                Kill TempFolderPath & TempPresentationName & "*.*"
            Else
                Exit Sub
            End If
        Case vbCancel
            Cancel = 1
        End Select
    Else
        TempPresentation.Close
        Kill TempFolderPath & TempPresentationName & "*.*"
    End If
End If

End Sub

Private Sub UserForm_Terminate()
MacrosPresentation.Close
End Sub
