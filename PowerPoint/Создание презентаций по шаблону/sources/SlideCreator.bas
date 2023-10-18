Attribute VB_Name = "SlideCreator"
Option Private Module

Public Const sVersion As String = "Создание презентаций. Версия 1.1 от 18/10/2023"

Type SlideDataType
    Template As Integer
    Text1 As String
    Text2 As String
    ImgFolderPath As String
    ImgFileName As String
    ImgFullPath As String
    imgW As Integer
    imgH As Integer
End Type

Public SlideData As SlideDataType
Public emptySlideData As SlideDataType

Public TempFolderPath As String
Public TempPresentationName As String
Public TempPresentation As Presentation
Public MacrosPresentation As Presentation
Public SlideArray() As String
Public oImage As Object

Sub ResetSlideData()
SlideData = emptySlideData


If UserForm.frmSlideText1.Visible Then
    SlideData.Text1 = UserForm.txtSlideText1.Text
End If

If UserForm.frmSlideText2.Visible Then
    SlideData.Text2 = UserForm.txtSlideText2.Text
End If

If UserForm.frmSlideImage.Visible Then
    If UserForm.lstFilesList.Value <> "" Then
        SlideData.ImgFullPath = UserForm.lstFilesList.Value
        UserForm.imgSlideImage.Picture = ThumbnailPicFromFile(SlideData.ImgFullPath, 256, 256)
        Set oImage = CreateObject("WIA.ImageFile")
        oImage.LoadFile SlideData.ImgFullPath
        SlideData.imgH = oImage.Height
        SlideData.imgW = oImage.Width
    End If
End If

If UserForm.optTemplate6 Then
    SlideData.Template = 6
ElseIf UserForm.optTemplate7 Then
    SlideData.Template = 7
ElseIf UserForm.optTemplate8 Then
    SlideData.Template = 8
ElseIf UserForm.optTemplate4 Then
    SlideData.Template = 4
ElseIf UserForm.optTemplate3 Then
    SlideData.Template = 3
ElseIf UserForm.optTemplate2 Then
    SlideData.Template = 2
ElseIf UserForm.optTemplate1 Then
    SlideData.Template = 1
ElseIf UserForm.optTemplate5 Then
    SlideData.Template = 5
End If

End Sub
Sub CreateSlide()
Dim oSlide As Slide
Dim iH As Integer
Dim iW As Integer
Dim iT As Integer
Dim iL As Integer

Let TempFolderPath = Environ("TMP") & "\"
If TempPresentation Is Nothing Then
    If TempPresentationName = "" Then Let TempPresentationName = "Новая_презентация_" & Format(Now, "YYYYMMDDHHNNSS")
    MacrosPresentation.SaveCopyAs TempFolderPath & TempPresentationName & ".pptx"
    Set TempPresentation = Presentations.Open(TempFolderPath & TempPresentationName & ".pptx", WithWindow:=False, ReadOnly:=msoFalse)
    For Each oSlide In TempPresentation.Slides
        oSlide.Delete
    Next oSlide
End If

Set oSlide = TempPresentation.Slides.AddSlide(TempPresentation.Slides.Count + 1, TempPresentation.SlideMaster.CustomLayouts(SlideData.Template))

ReDim Preserve SlideArray(TempPresentation.Slides.Count)
SlideArray(oSlide.SlideIndex) = TempFolderPath & TempPresentationName & "_" & oSlide.SlideIndex & ".jpg"

If SlideData.Text1 <> "" Then If oSlide.Shapes(1).HasTextFrame Then oSlide.Shapes(1).TextFrame.TextRange.Text = SlideData.Text1
If SlideData.Text2 <> "" Then If oSlide.Shapes(2).HasTextFrame Then oSlide.Shapes(2).TextFrame.TextRange.Text = SlideData.Text2
If SlideData.ImgFullPath <> "" Then
    If SlideData.imgW >= SlideData.imgH Then 'горизонтальная картинка или квадрат - ориентируемся на ширину
        iW = oSlide.Shapes(2).Width
        iH = oSlide.Shapes(2).Width / SlideData.imgW * SlideData.imgH
        If iH > oSlide.Shapes(2).Height Then 'если итоговая высота получилась больше высоты шейпа, пересчитываем
            iW = oSlide.Shapes(2).Height / iH * iW
            iH = oSlide.Shapes(2).Height
        End If
    ElseIf SlideData.imgW < SlideData.imgH Then 'вертикальная картинка - ориентируемся на высоту
        iH = oSlide.Shapes(2).Height
        iW = oSlide.Shapes(2).Height / SlideData.imgH * SlideData.imgW
        If iW > oSlide.Shapes(2).Width Then 'если итоговая ширина получилась больше ширины шейпа, пересчитываем
            iH = oSlide.Shapes(2).Width / iW * iH
            iW = oSlide.Shapes(2).Width
        End If
        
    End If
    oSlide.Shapes(2).Fill.UserPicture SlideData.ImgFullPath

    
    
    'oSlide.Shapes(2).Top = Abs(TempPresentation.PageSetup.SlideHeight - iH) / 2 ' + oSlide.Shapes(2).Top
    oSlide.Shapes(2).Top = Abs(oSlide.Shapes(2).Height - iH) / 2 + oSlide.Shapes(2).Top
    'oSlide.Shapes(2).Left =  Abs(oSlide.Shapes(2).Width - iW) / 2 + oSlide.Shapes(2).Left
    oSlide.Shapes(2).Left = Abs(oSlide.Shapes(2).Width - iW) / 2 + oSlide.Shapes(2).Left
    oSlide.Shapes(2).Width = iW
    oSlide.Shapes(2).Height = iH

End If

oSlide.Export SlideArray(oSlide.SlideIndex), "JPG", 256
'UserForm.



'TempPresentation.Save
'TempPresentation.Close

End Sub


Sub test()
Set MacrosPresentation = ActivePresentation
UserForm.Show
End Sub
