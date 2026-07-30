Attribute VB_Name = "SaveBitmapAs"
Option Explicit

Private GdipToken       As Long
Private GdipInitialized As Boolean

Public Const GdiPlusVersion     As Long = 1
Private Const CP_ACP            As Long = 0

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
   ClassID As GUID
   FormatID As GUID
   CodecName As Long
   DllName As Long
   FormatDescription As Long
   FilenameExtension As Long
   MimeType As Long
   flags As Long
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long
   SigMask As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As Long, ByVal hpal As Long, nBitmap As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal size As Long, ByRef Encoders As Any) As Long

Public Function SavePictureFromHDC(ByVal hBitmap As Long, ByVal sFileName As String) As Boolean
    Dim lBitmap As Long
    Dim PicEncoder As GUID
    Dim sID As String
    
    ' Use file name extention to determine,
    ' what format we want to save the file in.
    Select Case LCase$(Right$(sFileName, 4))
        Case ".png"
            sID = "image/png"
        Case ".gif"
            sID = "image/gif"
        Case ".jpg"
            sID = "image/jpeg"
        Case ".tif"
            sID = "image/tiff"
        Case ".bmp"
            sID = "image/bmp"
        Case Else
            Exit Function
    End Select
    
    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, lBitmap) = 0 Then
        If GetEncoderClsid(sID, PicEncoder) = True Then
            SavePictureFromHDC = (GdipSaveImageToFile(lBitmap, StrPtr(sFileName), PicEncoder, ByVal 0) = 0)
        End If
        GdipDisposeImage lBitmap
    End If
End Function

Private Function GetEncoderClsid(strMimeType As String, ClassID As GUID) As Boolean
    
    Dim num As Long
    Dim size As Long
    Dim imgCodecInfo() As ImageCodecInfo
    Dim lval As Long
    Dim buffer() As Byte

    GdipGetImageEncodersSize num, size
    If size Then
        ReDim imgCodecInfo(num) As ImageCodecInfo
        ReDim buffer(size) As Byte

        GdipGetImageEncoders num, size, buffer(0)
        CopyMemory imgCodecInfo(0), buffer(0), (Len(imgCodecInfo(0)) * num)

        For lval = 0 To num - 1
            'image/bmp,image/jpeg,image/gif,image/tiff,image/png
            If StrComp(GetStrFromPtrW(imgCodecInfo(lval).MimeType), strMimeType, vbTextCompare) = 0 Then
                ClassID = imgCodecInfo(lval).ClassID
                GetEncoderClsid = True
                Exit For
            End If
        Next
        Erase imgCodecInfo
        Erase buffer
    End If
    
End Function

Private Function GetStrFromPtrW(lpszW As Long) As String
    
    Dim sRV As String

    sRV = String$(lstrlenW(ByVal lpszW) * 2, vbNullChar)
    WideCharToMultiByte CP_ACP, 0, ByVal lpszW, -1, ByVal sRV, Len(sRV), 0, 0
    GetStrFromPtrW = Left$(sRV, lstrlenW(StrPtr(sRV)))
    
End Function

Public Sub StartUpGDIPlus(ByVal GdipVersion As Long)
    
    Dim GdipStartupInput As GDIPlusStartupInput
    
    GdipStartupInput.GdiPlusVersion = GdipVersion
    GdipInitialized = (GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0&) = 0)

End Sub

Public Sub ShutdownGDIPlus()
    
    If GdipInitialized Then
        GdiplusShutdown GdipToken
    End If
    
End Sub

