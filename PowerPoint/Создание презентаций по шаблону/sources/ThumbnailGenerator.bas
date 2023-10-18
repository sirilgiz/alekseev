Attribute VB_Name = "ThumbnailGenerator"
Option Explicit

Type Size
    cx As Long
    cy As Long
End Type

Private Type uPicDesc
    Size As Long
    Type As Long
    #If VBA7 Then
        hPic As LongPtr
        hPal As LongPtr
    #Else
        hPic As Long
        hPal As Long
    #End If
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, lpiid As Any) As Long
    Private Declare PtrSafe Function SHCreateItemFromParsingName Lib "shell32" (ByVal pPath As LongPtr, ByVal pBC As Long, rIID As Any, ppV As Any) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
#Else
    Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, lpiid As Any) As Long
    Private Declare Function SHCreateItemFromParsingName Lib "shell32" (ByVal pPath As Long, ByVal pBC As Long, rIID As Any, ppV As Any) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
#End If

Private Enum vtbl_IShellItemImageFactory
    QueryInterface
    AddRef
    Release
    GetImage
End Enum

#If Win64 Then
    Private Const VTBL_OFFSET = vtbl_IShellItemImageFactory.GetImage * 8
#Else
    Private Const VTBL_OFFSET = vtbl_IShellItemImageFactory.GetImage * 4
#End If


Private Const CC_STDCALL As Long = 4
Private Const S_OK = 0
Private Const vbPicTypeBitmap = 1
Private Const IID_IShellItemImageFactory = "{BCC18B79-BA16-442F-80C4-8A59C30C463B}"

Public Function ThumbnailPicFromFile(ByVal FilePath As String, Optional ByVal Width As Long = 32, Optional ByVal Height As Long = 32) As StdPicture

    #If Win64 Then
        Static hBmp As LongPtr
        Dim pUnk As LongPtr
        Dim lPt As LongPtr
    #Else
        Static hBmp As Long
        Dim pUnk As Long
    #End If

    Dim lRet As Long, bIID(0 To 15) As Byte, Unk As IUnknown
    Dim tSize As Size, sFilePath As String
    
    DeleteObject hBmp
    If Len(Dir(FilePath, vbDirectory)) Then
        If CLSIDFromString(StrPtr(IID_IShellItemImageFactory), bIID(0)) = S_OK Then
            If SHCreateItemFromParsingName(StrPtr(FilePath), 0, bIID(0), Unk) = S_OK Then
                pUnk = ObjPtr(Unk)
                If pUnk Then
                    tSize.cx = Width: tSize.cy = Height
                    #If Win64 Then
                        CopyMemory lPt, tSize, LenB(tSize)
                        If CallFunction_COM(pUnk, VTBL_OFFSET, vbLong, CC_STDCALL, lPt, 0, VarPtr(hBmp)) = S_OK Then
                    #Else
                        If CallFunction_COM(pUnk, VTBL_OFFSET, vbLong, CC_STDCALL, tSize.cx, tSize.cy, 0, VarPtr(hBmp)) = S_OK Then
                    #End If
                            If hBmp Then
                                Set ThumbnailPicFromFile = PicFromBmp(hBmp)
                            End If
                        End If
                End If
            End If
        End If
    End If
End Function


#If VBA7 Then
    Private Function PicFromBmp(ByVal hBmp As LongPtr) As StdPicture
    Dim hLib As LongPtr
#Else
    Private Function PicFromBmp(ByVal hBmp As Long) As StdPicture
    Dim hLib As Long
#End If
    
    Dim uPicDesc As uPicDesc, IID_IPicture As GUID, oPicture As IPicture
    
    With uPicDesc
        .Size = Len(uPicDesc)
        .Type = vbPicTypeBitmap
        .hPic = hBmp
        .hPal = 0
    End With
    
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    If OleCreatePictureIndirectAut(uPicDesc, IID_IPicture, True, oPicture) = S_OK Then
        Set PicFromBmp = oPicture
    End If

End Function


#If VBA7 Then
    Private Function CallFunction_COM(ByVal InterfacePointer As LongPtr, ByVal VTableOffset As Long, _
    ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    
    Dim vParamPtr() As LongPtr
#Else
    Private Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, _
    ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    
    Dim vParamPtr() As Long
#End If
    
    Dim vParamType() As Integer
    Dim pIndex As Long, pCount As Long
    Dim vRtn As Variant, vParams() As Variant

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function
    
    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If
                                                       
    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, vParamType(0), vParamPtr(0), vRtn)
        
    If pIndex = 0& Then
        CallFunction_COM = vRtn
    Else
        SetLastError pIndex
    End If
 
End Function


