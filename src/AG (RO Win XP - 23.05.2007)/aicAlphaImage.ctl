VERSION 5.00
Begin VB.UserControl aicAlphaImage 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MaskColor       =   &H80000014&
   PropertyPages   =   "aicAlphaImage.ctx":0000
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   ToolboxBitmap   =   "aicAlphaImage.ctx":0015
   Windowless      =   -1  'True
End
Attribute VB_Name = "aicAlphaImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Credits/Acknowledgements
'   Relies almost totally on my c32bppDIB Suite project. Credits included in that project
'       http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'   Paul Caton for his thunking routines.
'       Timer callbacks created using his code
' For most current updates/enhancements visit the following:
'   Visit http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1

' See the Usage.RTF file provided for more information

' Common Public Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object"
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object"
Public Event MouseExit()
Attribute MouseExit.VB_Description = "Occurs when the user first moves the mouse cursor out of the control"
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user first moves the mouse cursor into the control"
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus"
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus"
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual"
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual"

' Custom callbacks
Public Event FadeTerminated(ByVal CurrentOpacity As Long)
' ^^ when FadeInOut is called, this event is fired when the fade has completed
'    Note: It is not fired if the fade was terminated by you


'-Callback declarations for Paul Caton thunking magic----------------------------------------------
Private z_CbMem   As Long    'Callback allocated memory address
Private z_Cb()    As Long    'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

' Timer and HitTest related APIs
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

' Drawing related APIs
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetClipBox Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

' Window properties related APIs
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER As Long = 4

Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Type MASKUSAGE
    Color As Long           ' current mask color
    Applied As Boolean      ' mask has been applied
    AppliedColor As Long    ' color used to create mask; may not be same as Color
    Source As aiMaskSource  ' mask option: see aiMaskSource enum
End Type
Private Type FADERCONTROL
    tmrAddr As Long         ' AddressOf timer call back procedure
    fStep As Long           ' percent of opacity to change between steps
    fDelay As Long          ' length to delay before next step occurs
    fOpacity As Long        ' final opacity that also terminates the fader
End Type
Private Type SCALEDCOORD
    left As Long            ' position of image within usercontrol
    tOp As Long
    Width As Long           ' size of image within usercontrol
    Height As Long
    RotatedSize As Long     ' when rotated, the size needed to completely rotate image 360 degrees
    OneToOne As Boolean     ' flag used for painting; when image is actual size, faster renders
End Type

Public Enum aiMaskSource
    aiNoMask = 0
    aiUseMaskColor = 1
    aiUseTopLeft = 2
    aiUseTopRight = 3
    aiUseBottomLeft = 4
    aiUseBottomRight = 5
End Enum
Public Enum aiMirrorEnum
    aiMirrorNone = 0
    aiMirrorHorizontal = 1
    aiMirrorVertical = 2
    aiMirrorAll = 3
End Enum
Public Enum aiScaleMethod
    aiScaled = 0
    aiStretch = 1
    aiScaleDownOnly = 2
    aiActualSize = 3
    aiLockScale = 4
End Enum
Public Enum aiGrayScales
    aiNTSCPAL = 1     ' R=R*.299, G=G*.587, B=B*.114 - Default
    aiCCIR709 = 2     ' R=R*.213, G=G*.715, B=B*.072
    aiSimpleAvg = 3   ' R,G, and B = (R+G+B)/3
    aiRedMask = 4     ' uses only the Red sample value: RGB = Red / 3
    aiGreenMask = 5   ' uses only the Green sample value: RGB = Green / 3
    aiBlueMask = 6    ' uses only the Blue sample value: RGB = Blue / 3
    aiRedGreenMask = 7 ' uses Red & Green sample value: RGB = (Red+Green) / 2
    aiBlueGreenMask = 8 ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    aiNoGrayScale = 0
End Enum
Public Enum aiHitTestStyle  ' see HitTest property
    aiBoundingRgn = 1
    aiEnclosedRgn = 2
    aiShapedRgn = 3
    aiEntireControl = 0
End Enum
Public Enum aiOLEDropMode
    aiDropNone = vbOLEDropNone
    aiDropManual = vbOLEDropManual
End Enum

Private cKeyProps As Long
'1=HighQuality,2=Stretch,4=AutoSize,8=AutoRedraw;16=KeepBytes;32=OffscreenActive,64=MaskUsed

'//Rotation related variables
Private cRotated As Boolean
Private cAngle As Long

Private cHitTest As aiHitTestStyle
Private cRegion As Long     ' used when cHitTest is aiShapedRgn, aiEnclosedRgn
Private cRgnBox As RECT     ' used when cHitTest is aiEntireControl, aiBoundingRgn

Private cGrayScale As aiGrayScales
Private cScaleMode As ScaleModeConstants    ' parent container's scalemode; used for public events
Private cScaleMethod As aiScaleMethod
Private cMirror As aiMirrorEnum
Private cOpacity As Long
Private cMask As MASKUSAGE
Private cScaledCoords As SCALEDCOORD
Private cKeepOrigFormat As Boolean          ' testing. Still playing with ideas for this

Private cImage As c32bppDIB
Private cOffscreen As c32bppDIB             ' used when AutoRedraw=True

'//Timer & mouse enter/exit related variables
Private cProjOwner As Long
Private cPropKey As String
Private cTmrAddrOf As Long
Private cTmrHwnd As Long
Private cTopLeftPos As POINTAPI
Private cFader As FADERCONTROL

' This module is grouped/organized into the following sections:
'   Public Properties & Functions
'   Property Page Routines
'   Inter-Control Communication
'   Support/Helper Routines
'   Usercontrol Properties
'   Callback Thunking Routines
'=========================================================================================================
' /////////////////////// P U B L I C   P R O P E R T I E S  &  F U N C T I O N S \\\\\\\\\\\\\\\\\\\\\\\\
'=========================================================================================================

Public Sub ClearImage()
Attribute ClearImage.VB_Description = "Removes image from control"
    ' Purpose: Permanently remove the control's image source
    If ((cKeyProps And 32) = 32) Then Set cOffscreen = Nothing
    cImage.DestroyDIB
    If Ambient.UserMode = True Then
        UserControl.Refresh
    Else ' can be called by the property page during design time
        cImage.InitializeDIB cScaledCoords.Width, cScaledCoords.Height
        cImage.CreateCheckerBoard
        sptResize
    End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object"
    ' Purpose: Simply refresh the control
    UserControl.Refresh
End Sub

Public Function LoadImage_FromArray(inStream() As Byte, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromArray.VB_Description = "Option to load an image from a stream of data"
    ' Purpose. Load an image from array
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_Stream(inStream, desiredIconWidth, desiredIconHeight, , , cKeepOrigFormat, desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromArray = True
    End If
End Function
Public Function LoadImage_FromFile(FileName As String, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromFile.VB_Description = "Option to load an image from a file"
    ' Purpose. Load an image from a file
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_File(FileName, desiredIconWidth, desiredIconHeight, cKeepOrigFormat, desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromFile = True
    End If
End Function
Public Function LoadImage_FromStdPicture(stdPic As StdPicture) As Boolean
Attribute LoadImage_FromStdPicture.VB_Description = "Option to load an image from a standard picture object"
    ' Purpose. Load an image from stdPicture object
    If cImage.LoadPicture_StdPicture(stdPic) = True Then
        If cKeepOrigFormat Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
        End If
        sptReplaceImage
        LoadImage_FromStdPicture = True
    End If
End Function
Public Function LoadImage_FromClipboard() As Boolean
Attribute LoadImage_FromClipboard.VB_Description = "Option to load an image from the clipboard"
    ' Purpose. Load an image from the clipboard
    If cImage.LoadPicture_ClipBoard() Then
        If cKeepOrigFormat Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
            Erase srcData()
        End If
        sptReplaceImage
        LoadImage_FromClipboard = True
    End If
End Function
Public Function LoadImage_FromHandle(Handle As Long) As Boolean
Attribute LoadImage_FromHandle.VB_Description = "Option to load an image from an existing memory handle"
    ' Purpose. Load an image from an image handle
    If cImage.LoadPicture_ByHandle(Handle) Then
        If cKeepOrigFormat Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
        End If
        sptReplaceImage
        LoadImage_FromHandle = True
    End If
End Function
Public Function LoadImage_FromResource(VBglobal As IUnknown, ResSection As Variant, ResID As Variant, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromResource.VB_Description = "Option to load an image from a resource file"
    ' Purpose. Load an image from a resource file
    
    ' Pass VBglobal like so:  VB.Global
    ' Pass ResSection as: vbResBitmap, vbResIcon, "Custom", etc
    ' Pass ResID as:  101, "MyLogo", etc
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_Resource(ResID, ResSection, VBglobal, desiredIconWidth, desiredIconHeight, , , desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromResource = True
    End If
End Function
Public Function LoadImage_FromOrignalBytes(Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
    ' Purpose. Load an image from a file
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_FromOrignalFormat(desiredIconWidth, desiredIconHeight, desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromOrignalBytes = True
    End If
End Function

Public Function GetImageBytes(imgBytes() As Byte, ByRef scanWidth As Long, _
                                Optional ByVal asArray2D As Boolean = False, _
                                Optional ByVal asBGRformat As Boolean = True, _
                                Optional ByVal asBottomUp As Boolean = False, _
                                Optional ByVal asPremultiplied As Boolean = False) As Boolean
    
    ' Purpose: Return the DIB bytes in a 1 or 2 dimensional array
    ' Return value is false if no image exists
    
    ' [Parameters].
    ' imgBytes(): a dynamic byte array; will be redimensioned as needed
    ' scanWidth: placeholder for the scan width of the bytes returned
    ' asArray2D: if True, arrays is returned as (x,y) else array is 1 dimensional
    ' asBGRformat: if True, bytes are BGRA else bytes are RGBA
    ' asBottomUp: if True, first byte is bottom,left of image else is top,left of image
    ' asPremultiplied: if True, bytes are in premultiplied format else they are not
    
    If Not cImage.ImageType = imgCheckerBoard Then
        scanWidth = cImage.scanWidth
        GetImageBytes = cImage.GetDIBbits(imgBytes, asArray2D, asBGRformat, True, True, asBottomUp, , , , , asPremultiplied)
    End If
    
End Function
Public Function SetImageBytes(imgBytes() As Byte, Optional ByVal isArray2D As Boolean = False, _
                                Optional ByVal isBGRformat As Boolean = True, _
                                Optional ByVal isBottomUp As Boolean = False) As Boolean
    
    ' Purpose: Sets the DIB bytes with the passed modified bytes and refresh image
    ' Return value is True if the bytes were applied
    
    ' [Parameters].
    ' imgBytes(): a dynamic byte array; will be redimensioned as needed
    ' isArray2D: if True, arrays is returned as (x,y) else array is 1 dimensional
    ' isBGRformat: if True, bytes are BGRA else bytes are RGBA
    ' isBottomUp: if True, first byte is bottom,left of image else is top,left of image
    
    If cImage.SetDIBbits(imgBytes, isBGRformat, , , isBottomUp) = True Then
        If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
        UserControl.Refresh
        SetImageBytes = True
    End If

End Function

Public Sub GetImageScales(ByRef Width As Long, ByRef Height As Long, _
            Optional ByVal ScaleMethod As aiScaleMethod = -1, _
            Optional ByVal desiredWidth As Long, Optional ByVal desiredHeight As Long, _
            Optional ByVal asRotatedImage As Boolean = False)
            
    ' Purpose: Return the size of the image that will fit within the passed
    '       DC dimensions using the passed scale method
    ' Notes:
    '   1. Parameters are pixels
    '   2. To resize your control to allow full rotation, use Sqr(Width*Width+Height*Height)
    
    ' [Parameters]
    '   Width. Variable to hold the returned scaled width
    '   Height. Variable to hold the returned scaled height
    '   ScaleMethod. If -1 then the current scale method will be used else the passed method
    '   desiredWidth. The desired width. If not passed, the current usercontrol width is used
    '   desiredHeight.  The desired height. If not passed, the current usercontrol height is used
    '   asRotatedImage. Indicates the desired width/height allows full 360 degree rotation
    
    If cImage.Handle = 0& Then Exit Sub
            
    If ScaleMethod < aiScaled Then
        ScaleMethod = cScaleMethod
    ElseIf ScaleMethod > aiLockScale Then
        ScaleMethod = aiLockScale
    End If
    
    If ScaleMethod = aiLockScale Then
        Width = cScaledCoords.Width
        Height = cScaledCoords.Height
        Exit Sub
    End If
    
    If desiredWidth < 1 Then desiredWidth = UserControl.ScaleWidth
    If desiredHeight < 1 Then desiredHeight = UserControl.ScaleHeight
    
    Dim xRatio As Single, yRatio As Single
    Dim rotSize As Long
    
    If asRotatedImage Then
        If ScaleMethod = aiActualSize Then
            Width = cImage.Width
            Height = cImage.Height
        Else
            rotSize = Sqr(cImage.Width * cImage.Width + cImage.Height * cImage.Height)
            xRatio = cImage.Width / rotSize
            yRatio = cImage.Height / rotSize
            Select Case ScaleMethod
            Case aiStretch
                Width = desiredWidth * xRatio
                Height = desiredHeight * yRatio
            Case Else
                If desiredWidth > desiredHeight Then
                    rotSize = desiredHeight
                Else
                    rotSize = desiredWidth
                End If
                Width = rotSize * xRatio
                Height = rotSize * yRatio
                If ScaleMethod = aiScaleDownOnly Then
                    If Width > cImage.Width Or Height > cImage.Height Then
                        Width = cImage.Width
                        Height = cImage.Height
                    End If
                End If
            End Select
        End If
    Else
        Select Case ScaleMethod
        Case aiActualSize
            Width = cImage.Width
            Height = cImage.Height
        Case aiStretch
            Width = desiredWidth
            Height = desiredHeight
        Case Else
            xRatio = desiredWidth / cImage.Width
            yRatio = desiredHeight / cImage.Height
            If yRatio < xRatio Then xRatio = yRatio
            If ScaleMethod = aiScaleDownOnly Then
                If xRatio > 1 Then xRatio = 1
            End If
            Width = cImage.Width * xRatio
            Height = cImage.Height * xRatio
        End Select
    End If
    
End Sub
            

Public Sub FadeInOut(ByVal FinalOpacity As Long, Optional ByVal FadeStepPercent As Long = 10, Optional ByVal FadeDelayInterval As Long = 30)
    
    ' Purpose: Fade an image from one opacity to another at a stepped percentage while delaying n milliseconds between steps
    ' Note: all values should be positive numbers. Routines will double check and adjust as needed
    
    ' [Parameters]
    ' FinalOpacity: The opacity of the image when the fader terminates
    ' FadeStepPercent: What percentage of opaqueness image should change between intervals (valid values are 1 to 100)
    ' FadeDelayInterval: How long image should remain before next fade step occurs (valid values are 10 to max long value)
    
    ' Kill current fader if it exists
    If Not cFader.fStep = 0& Then
        KillTimer cTmrHwnd, -ObjPtr(Me)
        cFader.fStep = 0&
    End If
    
    ' validate final opacity parameters
    If FinalOpacity < 0& Then
        FinalOpacity = 0&
    ElseIf FinalOpacity > 100& Then
        FinalOpacity = 100&
    End If
    ' validate step value
    If FadeStepPercent = 0& Then
        Exit Sub
    Else
        ' set up fader control
        If Not cOpacity = FinalOpacity Then
            If FinalOpacity < cOpacity Then
                FadeStepPercent = -Abs(FadeStepPercent) ' must be negative
            Else
                FadeStepPercent = Abs(FadeStepPercent)  ' else must be positive
            End If
            With cFader
                ' get AddressOf for our fader timer
                If .tmrAddr = 0& Then .tmrAddr = zb_AddressOf(2, 4, 1)
                .fOpacity = FinalOpacity        ' set final opacity value
                .fStep = FadeStepPercent        ' set fader step value
                If FadeDelayInterval < 10 Then  ' set fader interval
                    .fDelay = 10
                Else
                    .fDelay = FadeDelayInterval
                End If
            End With
            ' fire first fader event
            Call Timer_Fader(cTmrHwnd, 0&, -ObjPtr(Me), 0&)
        End If
    End If
    
End Sub

Public Property Let AutoRedraw(Enable As Boolean)
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap"
    ' Purpose: create/maintain a scaled/drawn image in its own DC.
    ' Although it uses more resources, images are rendered much faster since they
    ' to not have to be resized, rotated, mirrored, etc, each time the control must be updated
    ' Additionally, since the cached image is now at a 1:1 ratio, rendering only occurs to
    ' the invalidated portion of the control, not the entire control each time
    If Not (((cKeyProps And 8) = 8) = Enable) Then
        cKeyProps = cKeyProps Xor 8
        If Ambient.UserMode = True Then
            cKeyProps = cKeyProps Xor 32
            If Enable = False Then Set cOffscreen = Nothing
        End If
        PropertyChanged "AutoRedraw"
    End If
End Property
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = ((cKeyProps And 8) = 8)
End Property

Public Property Let HitTest(Style As aiHitTestStyle)
Attribute HitTest.VB_Description = "Returns/Sets method used to determine whether control responds to mouse events"
    ' Determines whether the mouse in the control registers as within the image
    ' Possible values are:
    ' - aiEntireControl. Mouse anywhere within the control is allowed
    ' - aiShapedRgn. Mouse on any non-transparent pixel is allowed
    ' - aiBoundingRgn. Mouse within the tightest rectangle around image is valid
    ' - aiEnclosedRgn. Same as aiShapedRgn but any transparent pixels that fall
    '       between non-transparent pixels is considered non-transparent
    If Style >= aiEntireControl And Style <= aiShapedRgn Then
        If Not Style = cHitTest Then
            cHitTest = Style
            sptRefreshRegion
            PropertyChanged "HitTest"
        End If
    End If
End Property
Public Property Get HitTest() As aiHitTestStyle
    HitTest = cHitTest
End Property

Public Property Let MaskColor(Color As OLE_COLOR)
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the image"
    ' Allows making a color within a non-transparent bitmap transparent
    If cMask.Applied Then ' previous mask already applied, remove it
        If cMask.Source = aiUseMaskColor Then
            sptUndoMask
            If (cKeyProps And 64) = 64 Then
                cImage.MakeTransparent sptConvertSysColor(Color) Xor vbWhite
            Else
                cImage.MakeTransparent sptConvertSysColor(Color)
            End If
            cMask.AppliedColor = Color
            If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
            UserControl.Refresh
        End If
    End If
    cMask.Color = Color
    PropertyChanged "MaskColor"
End Property
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = cMask.Color
End Property

Public Property Let MaskUsed(Style As aiMaskSource)
Attribute MaskUsed.VB_Description = "Returns/Sets whether the mask is to be applied to the image"
    ' Enables the mask to take effect.
    ' Only applies to non-transparent images
    If cMask.Applied = False And cImage.Alpha = True Then Exit Property
    If Style >= aiNoMask And Style <= aiUseBottomRight Then
        If Not cMask.Source = Style Then
            If Not cImage.ImageType = imgCheckerBoard Then
                Dim mColor As Long
                If cMask.Applied Then sptUndoMask
                If Style = aiNoMask Then
                    cMask.Applied = False
                Else
                    Select Case Style
                    Case aiUseMaskColor
                        cMask.AppliedColor = cMask.Color
                    Case aiUseBottomLeft
                        cMask.AppliedColor = cImage.GetPixel(0, cImage.Height - 1, , False)
                    Case aiUseBottomRight
                        cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, cImage.Height - 1, , False)
                    Case aiUseTopLeft
                        cMask.AppliedColor = cImage.GetPixel(0, 0, , False)
                    Case aiUseTopRight
                        cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, 0, , False)
                    End Select
                    cMask.Applied = True
                    If (cKeyProps And 64) = 64 Then
                        If Style = aiUseMaskColor Then
                            mColor = sptConvertSysColor(cMask.AppliedColor) Xor vbWhite
                        Else
                            mColor = cMask.AppliedColor
                            cMask.AppliedColor = mColor Xor vbWhite
                        End If
                    Else
                        mColor = sptConvertSysColor(cMask.AppliedColor)
                    End If
                    cImage.MakeTransparent mColor
                End If
                If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
                UserControl.Refresh
                cMask.Source = Style
                PropertyChanged "MaskUsed"
            End If
        End If
    End If
End Property
Public Property Get MaskUsed() As aiMaskSource
    MaskUsed = cMask.Source
End Property

Public Property Let InversedImage(Inverse As Boolean)
Attribute InversedImage.VB_Description = "Returns/Sets whether the image colors are inverted"
    ' Purpose: Inverts the colors of an image. Can be toggled
    If Not (((cKeyProps And 64) = 64) = Inverse) Then
        cKeyProps = cKeyProps Xor 64
        cImage.MakeImageInverse
        PropertyChanged "InversedImage"
        sptUpdateOffscreen False, False
        UserControl.Refresh
    End If
End Property
Public Property Get InversedImage() As Boolean
    InversedImage = ((cKeyProps And 64) = 64)
End Property

Public Property Let AutoSize(Value As Boolean)
    ' Purpose: Forces control to resize to the scaled image size
    If Not (((cKeyProps And 4) = 4) = Value) Then
        cKeyProps = cKeyProps Xor 4
        sptResize
        PropertyChanged "AutoSize"
    End If
End Property
Public Property Get AutoSize() As Boolean
    AutoSize = ((cKeyProps And 4) = 4)
End Property

Public Property Let ScaleMethod(Method As aiScaleMethod)
    ' Purpose: Determine whether image is scaled proportionally or not
    If Not Method = cScaleMethod Then
        cScaleMethod = Method
        Call sptResize
        PropertyChanged "Stretch"
    End If
End Property
Public Property Get ScaleMethod() As aiScaleMethod
    ScaleMethod = cScaleMethod
End Property

Public Property Let StretchQuality(highQuality As Boolean)
Attribute StretchQuality.VB_Description = "Returns/sets whether a graphic will be resized using the best sizing algorithms"
    ' Determines algoritm to use when stretching an image
    If Not (((cKeyProps And 1) = 1) = highQuality) Then
        cKeyProps = cKeyProps Xor 1
        cImage.HighQualityInterpolation = highQuality
        UserControl.Refresh
        PropertyChanged "StretchQuality"
    End If
End Property
Public Property Get StretchQuality() As Boolean
    StretchQuality = ((cKeyProps And 1) = 1)
End Property

Public Property Let Opacity(ByVal Opaqueness As Long)
Attribute Opacity.VB_Description = "Returns/Sets the level of translucency for the control. 100 is fully opaque and 0 is transparent"
    ' Purpose: Set how opaque the image will be rendered at; 100=fully opaque, 0=fully transparent
    If Not Opaqueness = 0 Then
        Opaqueness = Abs(Opaqueness) Mod 100
        If Opaqueness = 0 Then Opaqueness = 100
    End If
    cOpacity = Opaqueness
    PropertyChanged "Opacity"
    UserControl.Refresh
End Property
Public Property Get Opacity() As Long
    Opacity = cOpacity
End Property

Public Property Let Mirror(MirrorType As aiMirrorEnum)
Attribute Mirror.VB_Description = "Returns/Sets the current mirroring effect for the image"
    ' Purpose. Mirror an image horizontally or vertically or both
    If MirrorType >= aiMirrorNone And MirrorType <= aiMirrorAll Then
        sptMirrorImage MirrorType
        cMirror = MirrorType
        If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
        UserControl.Refresh
        PropertyChanged "Mirror"
    End If
End Property
Public Property Get Mirror() As aiMirrorEnum
    Mirror = cMirror
End Property

Public Property Let KeepOriginalFormat(bValue As Boolean)
Attribute KeepOriginalFormat.VB_Description = "Returns/Sets whether control will maintain original image data"
    ' Forces routines to keep a copy of the image in its original format
    ' This will be used more in upcoming versions
    If Ambient.UserMode = False Then cKeepOrigFormat = bValue
    PropertyChanged "KeepOriginalFormat"
End Property
Public Property Get KeepOriginalFormat() As Boolean
    KeepOriginalFormat = cKeepOrigFormat
End Property

Public Property Let grayScale(Style As aiGrayScales)
Attribute grayScale.VB_Description = "Returns/Sets gray scale formula used when rendering image"
    ' Option to toggle grayscale effect
    ' Note: offsetting -1/+1 to align values with the c32bppDIB enumeration
    ' That enumeration does not have a "No Grayscale" option
    If Style >= aiNoGrayScale And Style <= aiBlueGreenMask Then
        If Not cGrayScale = Style - 1 Then
            cGrayScale = Style - 1
            UserControl.Refresh
            PropertyChanged "GrayScale"
        End If
    End If
End Property
Public Property Get grayScale() As aiGrayScales
    grayScale = cGrayScale + 1
End Property

Public Property Let Rotation(ByVal newAngle As Long)
    ' Purpose: Rotate an image by set a degree (-360 to 360)
    ' See also RotateZeroDeg property for images at zero degrees
    newAngle = newAngle Mod 360
    If Not newAngle = cAngle Then
        If cAngle = 0& Then
            cAngle = newAngle
            cRotated = True
            sptResize
        Else
            cAngle = newAngle
            cRotated = Not (cAngle = 0&)
            If cAngle = 0& And ((cKeyProps And 4) = 4) Then
                sptResize
            Else
                If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
                UserControl.Refresh
            End If
        End If
        PropertyChanged "Rotation"
    End If

End Property
Public Property Get Rotation() As Long
    Rotation = cAngle
End Property

Public Property Let Enabled(Enable As Boolean)
    ' Purpose: Enable/Disable mouse events
    If Not UserControl.Enabled = Enable Then
        UserControl.Enabled = Enable
        If Ambient.UserMode = True Then
            If Enabled = True Then
                sptValidateSession
            Else
                sptInvalidateSession
            End If
            sptRefreshRegion
        End If
        PropertyChanged "Enabled"
    End If
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let MousePointer(Pointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object"
    ' same as VB's MousePointer property
    On Error Resume Next
    UserControl.MousePointer = Pointer
    If Err Then Err.Clear
    PropertyChanged "MousePointer"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MouseIcon(Icon As StdPicture)
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon"
    ' same as VB's MouseIcon property
    Set Me.MouseIcon = Icon
End Property
Public Property Set MouseIcon(Icon As StdPicture)
    On Error Resume Next
    Set UserControl.MouseIcon = Icon
    If Err Then Err.Clear
    PropertyChanged "MouseIcon"
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let OLEDropMode(Value As aiOLEDropMode)
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target"
    ' same as VB's OLEDropMode property
    ' Will forward the OLEDragOver, OLEDragDrop when set
    ' You can use this to allow the control to load an image dragged onto it
    On Error Resume Next
    UserControl.OLEDropMode = Value
    If Err Then Err.Clear
    PropertyChanged "OLEDropMode"
End Property
Public Property Get OLEDropMode() As aiOLEDropMode
    OLEDropMode = UserControl.OLEDropMode
End Property



'=============================================================================================
' /////////////////////// P R O P E R T Y   P A G E   R O U T I N E S \\\\\\\\\\\\\\\\\\\\\\\\
'=============================================================================================

Friend Function ppgGetStream(outStream() As Byte) As Boolean
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows the property page to retrieve this control's image remotely
    If Not cImage.ImageType = imgCheckerBoard Then ppgGetStream = cImage.GetOrginalFormat(outStream)
End Function
Friend Sub ppgSetStream(inStream() As Byte, cX As Long, cY As Long, bitDepth As Long)
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows the property page to set this control's new iamge remotely & triggers an activation of the WriteProperties event
    Dim curScale As aiScaleMethod, bAutoSize As Boolean
    cImage.LoadPicture_Stream inStream(), cX, cY, , , True, bitDepth
    curScale = Me.ScaleMethod
    bAutoSize = Me.AutoSize
    cKeyProps = (cKeyProps Or 4) ' turn autosize on
    cScaleMethod = aiActualSize      ' set scalemode to actual size
    sptResize
    cScaleMethod = curScale
    If Not bAutoSize Then cKeyProps = cKeyProps Xor 4
    PropertyChanged "ScaleMethod"
End Sub
Friend Property Get ppgDIBclass() As c32bppDIB
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows property page to access this controls DIB classes
    Set ppgDIBclass = cImage
End Property

'=============================================================================================
' ////////////////// I N T E R - C O N T R O L   C O M M U N I C A T I O N \\\\\\\\\\\\\\\\\\\
'=============================================================================================

Friend Sub iccRemoteMouseExit()
    ' When this control thinks it has the mouse over it, and another
    ' control receives a mouse event before the timer fires, that
    ' control will forward a mouse exit via this routine.
    If UserControl.Tag = "Timer" Then
        KillTimer cTmrHwnd, ObjPtr(Me)
        UserControl.Tag = vbNullString
    End If
    RaiseEvent MouseExit
End Sub


'===============================================================================================
' /////////////////////// S U P P O R T / H E L P E R   R O U T I N E S \\\\\\\\\\\\\\\\\\\\\\\\
'===============================================================================================


Private Sub sptReplaceImage()
    ' Function replaces an image and resizes the control
    ' But it also ensures current settings are applied if applicable and
    ' resets those that cannot apply
    sptMirrorImage 0&       ' mirror the image using current settings
    If cMask.Applied Then   ' did we apply a mask?
        If cImage.Alpha = True Then ' if the image is alpha, then masking is not applicable
            cMask.Applied = False   ' reset key masking values
            cMask.Source = aiNoMask
        Else
            Select Case cMask.Source    ' apply the mask to this image
                Case aiUseBottomLeft
                    cMask.AppliedColor = cImage.GetPixel(0, cImage.Height - 1, , False)
                Case aiUseBottomRight
                    cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, cImage.Height - 1, , False)
                Case aiUseTopLeft
                    cMask.AppliedColor = cImage.GetPixel(0, 0, , False)
                Case aiUseTopRight
                    cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, 0, , False)
            End Select
            cImage.MakeTransparent sptConvertSysColor(cMask.AppliedColor)
        End If
    End If
    If (cKeyProps And 64) = 64 Then cImage.MakeImageInverse
    sptResize

End Sub

Private Sub sptRefreshRegion()
    ' Routine creates a new hit test region; can be called for many reasons:
    ' - image is mirrored, changing shape of image
    ' - image is resized
    ' - image toggles mask property
    ' - image is rotated
    ' - hit test property changed
    ' - enabled property toggled
    If UserControl.Enabled = True Then
        If Ambient.UserMode = True Then
            If cHitTest = aiEntireControl Or cImage.Handle = 0& Then ' no region needed
                If Not cRegion = 0& Then
                    DeleteObject cRegion
                    cRegion = 0&
                End If
                SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
            Else
                If Not cRegion = 0& Then DeleteObject cRegion
                If Not cOffscreen Is Nothing Then
                    ' we have offscreen image, use it to create region
                    cRegion = cOffscreen.CreateRegion(cHitTest - 1)
                Else ' we are going to force offscreen to create the region
                    cKeyProps = cKeyProps Or 32
                    sptUpdateOffscreen True, False
                    cRegion = cOffscreen.CreateRegion(cHitTest - 1)
                    cKeyProps = cKeyProps Xor 32
                    Set cOffscreen = Nothing
                End If
                If cHitTest = aiBoundingRgn Then
                    ' we don't need to use a GDI resource, so don't
                    GetRgnBox cRegion, cRgnBox
                    DeleteObject cRegion
                    cRegion = 0&
                End If
            End If
        Else    ' design mode, use complete control
            SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
    Else    ' not used if in runtime, but used if in design time
        SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
End Sub

Private Function sptConvertSysColor(Color As Long) As Long

    ' Converts VB color constants to real color values
    If Color < 0 Then
        sptConvertSysColor = GetSysColor(Color And &HFF&)
    Else
        sptConvertSysColor = Color
    End If
    
End Function

Private Function sptUpdateOffscreen(bResize As Boolean, bUpdateRegion As Boolean) As Boolean
    ' Purpose: Maintain an offscreen image when the user has set the HasDC property to true
    
    If (cKeyProps And 32) = 32 Then
        
        Dim hdc As Long
        If bResize = True Or cOffscreen Is Nothing Then
            If cOffscreen Is Nothing Then
                Set cOffscreen = New c32bppDIB
                cOffscreen.ManageOwnDC = True
                cOffscreen.isAlphaBlendFriendly = True
            End If
            If cScaledCoords.Width = 0 Then Exit Function
            cOffscreen.InitializeDIB UserControl.ScaleWidth, UserControl.ScaleHeight
        Else
            cOffscreen.EraseDIB
        End If
        
        hdc = cOffscreen.LoadDIBinDC(True)
        If Not hdc = 0 Then
            If cRotated Then
                cOffscreen.Alpha = True
                cImage.RotateAtCenterPoint hdc, cAngle, cOffscreen.Width \ 2, cOffscreen.Height \ 2, cScaledCoords.Width, cScaledCoords.Height, , , , , , cOffscreen
            Else
                cOffscreen.Alpha = cImage.Alpha
                cImage.Render hdc, cScaledCoords.left, cScaledCoords.tOp, cScaledCoords.Width, cScaledCoords.Height, , , , , , , False, cOffscreen
            End If
        End If
        If bUpdateRegion Then sptRefreshRegion
        sptUpdateOffscreen = True
    End If

End Function

Private Sub sptResize()
    ' Purpose: Resize a control but monitor whether or not resizing actually occurred
    ' When we call the resize event which we want to do so the cScaleWidth/Height
    ' variables are recalculated, no repainting will occur if the image is already the
    ' correct size. So monitor and refresh in that case
    Dim cX As Long, cY As Long
    cX = UserControl.ScaleWidth
    cY = UserControl.ScaleHeight
    Call UserControl_Resize
    If UserControl.ScaleWidth = cX Then
        If UserControl.ScaleHeight = cY Then UserControl.Refresh
    End If
End Sub

Private Sub sptUndoMask()
    ' Routine replaces the applied mask color with the original color when the
    ' user toggled the MaskUsed property to false
    Dim X As Long, lColor As Long, srcData() As Byte
    
    ' get the masked color we applied to the image
    If (cKeyProps And 64) = 64 Then
        lColor = sptConvertSysColor(cMask.AppliedColor) Xor vbWhite
    Else
        lColor = sptConvertSysColor(cMask.AppliedColor)
    End If
    ' convert it to BGR and add an alpha value of 255
    lColor = (lColor And &HFF) * &H10000 Or ((lColor \ &H100) And &HFF) * &H100 _
        Or ((lColor \ &H10000) And &HFF) Or &HFF000000
    ' return the RGBA in a 1D array
    cImage.GetDIBbits srcData, False
    ' loop thru. Only those that are fully transparent were changed; reset them
    For X = 3 To UBound(srcData) Step 4&
        If srcData(X) = 0 Then
            CopyMemory srcData(X - 3), lColor, 4&
        End If
    Next
    ' apply the changed RGBA bytes
    cImage.SetDIBbits srcData
End Sub

Private Sub sptMirrorImage(newMirrorValue As aiMirrorEnum)
    ' Purpose: Permanently mirror an image.
    ' Mirroring this way improves rendering speed vs mirroing during rendering.
    ' Unrendering is just as simple and mirroring does not destroy any pixel information
    Dim curMirror As aiMirrorEnum
    curMirror = (cMirror Xor newMirrorValue)
    cImage.MirrorImage ((curMirror And aiMirrorHorizontal) = aiMirrorHorizontal), ((curMirror And aiMirrorVertical) = aiMirrorVertical)
End Sub

Private Sub sptValidateSession()
    
    If Ambient.UserMode = False Then Exit Sub
    
    ' The sptValidateSession and sptInvalidateSession are necessary only when playing with
    ' the controls, uncompiled.  But since there is no guarantee you would take my
    ' advice and compile this OCX vs leaving uncompiled within your project, I
    ' opted to add these safety procedures.
    
    ' Feedback between controls when they are compiled can occur using mapped memory files,
    ' mutexes, window properties, subclassing and other strategies. When compiled, this is
    ' rather easy because when the application terminates, whether or not END was executed,
    ' the controls get their terminate event and can clean up and zeroize references.
    ' But when uncompiled, and END is executed, terminate events do not occur. Well if user
    ' hits Stop/End then we can't clean up our references/resources while in IDE: we still
    ' have a mapped file, we still have a mutex, etc, etc. If we are storing references in
    ' those objects and the objects aren't destroyed but the things that the references point
    ' to are, then when we try to access them next time the IDE project is run, we get a
    ' memory violation and crash. So we need a good system for the uncompiled controls to know
    ' whether or not they are running fresh every time F5 is pressed in IDE. This way we can
    ' know whether any existing references are valid or not.
    
    ' In order to provide MouseEnter/MouseExit feedback, I need a way to have each
    ' control talk to one another so one can terminate the faked mouse capture
    ' when another takes it over. Remember, these are windowless so we can't
    ' use SetCapture and TrackMouseEvent like normal windowed controls. We know if
    ' the mouse is in our control via the HitTest event, but we won't know when
    ' it leaves our control without some creative workarounds. Disabled controls
    ' receive no mouse events and no HitTest events
    
    ' Anyway, I have a couple of choices, one that involves hooking the process and
    ' and another that will use a timer. The timer method I chose for safety only.
    ' However, whether hooking (API call backs) or timer method (creating soft
    ' references to controls), I have to make absolutely sure we don't try to
    ' access memory not ours which can, and most likely would, happen when uncompiled
    ' controls are run in IDE and the user hits the END or STOP button since
    ' the controls never fire their terminate event while in IDE and can't clean up.
    
    ' So, to prevent this from crashing your apps, I am using window properties
    ' to store information. But that information must be correct after STOP or END
    ' was executed, so I need a way to know so that I can erase the properties
    ' and start fresh, just like the app will be doing.  Can't use window handles
    ' nor window procedures, nor thread IDs because they do not always change from
    ' runtime to runtime -- I've tried. Therefore, to be able to know if you hit
    ' END or STOP, I use a cross-reference system and reference count like so.
    
    ' Note: The hidden VB owner window, while in IDE, never closes until VB is closed.
    ' Each time a control is displayed that is not disabled, we find the control's
    ' top-level parent form and then find the hidden owner of that form. On that owner,
    ' we increment a reference count against the top-level parent's hWnd and on the
    ' parent, we add a reference to the owner. Then when the controls are destroyed
    ' normally (not END or STOP), the reverse occurs... Reference counts are decremented
    ' and when they get to zero, the parent's reference to the owner is removed.
    
    ' Now when you start the app, this routine checks to see if the owner has any parent
    ' references, if they do then we verify the references are actual windows. If not,
    ' STOP or END happend. If that is good to go, we double check (cross-referenced) to
    ' see if the parent has a reference to the owner (newly created windows won't have
    ' any properties to read). If not, then STOP or END happend. Bottom line, the owner
    ' and parent either both must have references to each other or neither do, otherwise
    ' the controls were not destroyed normally.
    
    ' The above may sound like gibberish or you may not quite get it.  All I can say, is
    ' imagine you are creating a usercontrol that can exist both compiled (gets terminate
    ' events) and uncompiled (does not get terminate events), that they can also be placed
    ' on multiple containers in multiple forms and that each of those controls needs a way
    ' to talk to each other. Now imagine that if the data shared between the controls was
    ' not correct your app will crash. That is the scenario.
    

    Dim lValue As Long
    Dim pHwnd As Long
    Dim w As Long, pOwner As Long
    Dim bPurge As Boolean
    Dim bReferenced As Boolean
    
    If cTmrAddrOf = 0& Then cTmrAddrOf = zb_AddressOf(1, 4)
    
    ' get top level parent window
    pHwnd = UserControl.ContainerHwnd
    cTmrHwnd = pHwnd
    Do
        lValue = GetParent(pHwnd)
        If lValue = 0& Then Exit Do
        pHwnd = lValue
    Loop
    cProjOwner = GetWindow(pHwnd, GW_OWNER)
    If cTmrHwnd = 0& Then cTmrHwnd = pHwnd
    
    If cProjOwner = 0& Then
        cProjOwner = pHwnd  ' won't happen in VB; could happen in other containers (IE maybe?)
    Else
        lValue = cProjOwner ' is this the top level Owner?
        Do
            lValue = GetWindow(cProjOwner, GW_OWNER)
            If lValue = 0 Then Exit Do
            cProjOwner = lValue
        Loop
    End If
    
    ' The hidden VB owner will have a key for each form: Client0, Client1, etc
    lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
    Do Until lValue = 0&
        ' ensure the current hWnd references cProjOwner in its property
        If GetProp(lValue, cPropKey & "Parent") = cProjOwner Then
            ' it is, has this control's parent been counted?
            If lValue = pHwnd Then bReferenced = True
        Else
            ' invalid references we purge current properties, if any
            bPurge = True
            Exit Do
        End If
        
        w = w + 1
        lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
    Loop
            
    If bPurge Then
        Debug.Print "Purging data"
        For w = 0 To w
            RemoveProp cProjOwner, cPropKey & "Client" & w
        Next
        lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
        Do Until lValue = 0&
            RemoveProp cProjOwner, cPropKey & "Client" & w
            w = w + 1
            lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
        Loop
        ' this is important: if "Capture" property is invalid, we could try to
        ' iccRemoteMouseExit on an invalid object which would cause a GPF.
        ' This is the entire reason for the Validate/InvalidateSession routines
        RemoveProp cProjOwner, cPropKey & "Capture"
        bReferenced = False
        w = 0
    End If
    If Not bReferenced Then
        SetProp cProjOwner, cPropKey & "Client" & w, pHwnd
        SetProp pHwnd, cPropKey & "Parent", cProjOwner
        SetProp cProjOwner, cPropKey & "Ref" & pHwnd, 1
        'Debug.Print "setting new reference count "; pHwnd; 1
    Else
        lValue = GetProp(cProjOwner, cPropKey & "Ref" & pHwnd)
        SetProp cProjOwner, cPropKey & "Ref" & pHwnd, lValue + 1
        'Debug.Print "incrementing "; pHwnd, lValue + 1
    End If
End Sub

Private Sub sptInvalidateSession()

    ' This is the complimentary function of ValidateSesssion -- See that routine for comments
    If Ambient.UserMode = False Then Exit Sub

    Dim w As Long
    Dim pHwnd As Long
    Dim lValue As Long
    Dim propNr As Long, refCount As Long
    
    ' kill timer and clean up timer data if needed
    If Not cTmrAddrOf = 0& Then
        If UserControl.Tag = "Timer" Then
            KillTimer cTmrHwnd, ObjPtr(Me)
            UserControl.Tag = vbNullString
        End If
        zTerminate
    End If
    ' ensure tghe Capture property does not reference this control
    If GetProp(cProjOwner, cPropKey & "Capture") = ObjPtr(Me) Then
        SetProp cProjOwner, cPropKey & "Capture", 0&
    End If
    
    ' get the top-level parent window
    pHwnd = cTmrHwnd
    Do
        lValue = GetParent(pHwnd)
        If lValue = 0& Then Exit Do
        pHwnd = lValue
    Loop
    
    ' find this control's top-level parent in the project owner's properties
    lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
    Do Until lValue = 0&
        If lValue = pHwnd Then
            propNr = w
            refCount = GetProp(cProjOwner, cPropKey & "Ref" & pHwnd) - 1
        End If
        w = w + 1
        lValue = GetProp(cProjOwner, cPropKey & "Client" & w)
    Loop
    If refCount < 1& Then
        'Debug.Print "Removing refs "; pHwnd, "owner is "; cProjOwner
        Select Case w - 1
        Case propNr ' this control's pHwnd is the last in the property listing. Is it also the only one listed?
            If propNr = 0 Then ' this is the last referenced control; remove the capture property too
                ' this is important: if it remains, we could try to iccRemoteMouseExit on an invalid object
                ' which would cause a GPF. This is the entire reason for the Validate/InvalidateSession routines
                RemoveProp cProjOwner, cPropKey & "Capture"
                Debug.Print "Removed Capture references"
            End If
        Case Is > propNr ' Move the last pHwnd in the property list to this pHwnd's position
            SetProp cProjOwner, cPropKey & "Client" & propNr, GetProp(cProjOwner, cPropKey & "Client" & w - 1)
            propNr = w - 1& ' re-reference so the last item is removed from the property list
        End Select
        RemoveProp cProjOwner, cPropKey & "Client" & propNr ' remove the pHwnd from the list
        RemoveProp pHwnd, cPropKey & "Parent"               ' remove the project owner reference from this pHwnd
    Else
        SetProp cProjOwner, cPropKey & "Ref" & pHwnd, refCount  ' decrement number of controls on this pHwnd
        'Debug.Print "Decrementing count "; pHwnd; refCount
    End If
    
End Sub

'============================================================================================
' /////////////////////// U S E R C O N T R O L  P R O P E R T I E S \\\\\\\\\\\\\\\\\\\\\\\\
'============================================================================================

Private Sub UserControl_Initialize()
    Set cImage = New c32bppDIB
    cImage.isAlphaBlendFriendly = True
    cPropKey = "AIC" & App.Major & "." & App.Minor & ":"
End Sub

Private Sub UserControl_Terminate()
    Set cOffscreen = Nothing    ' clean up any offscreen image
    If Not cRegion = 0& Then DeleteObject cRegion
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(y, vbPixels, cScaleMode), State)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Hide()
    ' control is made invisible by user or is about to be destroyed
    If Not cFader.fStep = 0& Then
        ' stop fader timer
        KillTimer cTmrHwnd, -ObjPtr(Me)
        cFader.fStep = 0&
    End If
    ' stop mouse_exit timer & clean up references as needed
    If UserControl.Enabled Then sptInvalidateSession
End Sub

Private Sub UserControl_Show()

    If UserControl.Enabled Then sptValidateSession
    If cScaledCoords.Width = 0& Then sptResize
    ' Thanx to Soorya for helping trouble shoot above issue.
    ' A usercontrol while in IDE will not always fire a Resize event before it is
    ' first displayed. Since scaling ratios are calculated based off the usercontrol's
    ' size and Stretch property, we need it fired, so fire it here if not previously
    ' fired; otherwise images are not drawn correctly or not drawn at all!
End Sub

Private Sub UserControl_HitTest(X As Single, y As Single, HitResult As Integer)

    ' For windowless, DC-less controls, this is the only way to let VB know
    ' whether or not we want the mouse or let it go to next object in zOrder.
    ' Note: This is used both during design time and run time
    
    If cRegion = 0& Then
        If Not PtInRect(cRgnBox, X, y) = 0& Then HitResult = vbHitResultHit
    ElseIf Not PtInRegion(cRegion, X, y) = 0& Then
        HitResult = vbHitResultHit
    End If
    
    If Ambient.UserMode = True Then
        ' testing for mouse enter/mouse leave
        ' See sptValidateSession and sptInvalidateSession regarding these window properties
        
        Dim hCapturedAIC As Long, myPointer As Long, mPoint As POINTAPI
        
        ' get object pointer of control currently having mouse enter
        hCapturedAIC = GetProp(cProjOwner, cPropKey & "Capture")
        myPointer = ObjPtr(Me)
        
        Select Case hCapturedAIC
        Case myPointer ' same control; no change?
            If HitResult = 0& Then  ' we are no longer in the hit region
                ' trigger a mouse leave event
                KillTimer cTmrHwnd, myPointer
                UserControl.Tag = vbNullString
                SetProp cProjOwner, cPropKey & "Capture", 0&
                RaiseEvent MouseExit
            End If
        Case 0& ' no alphaImage control has the mouse, but we may now
            If HitResult = vbHitResultHit Then
                ' convert screen coordinates to relative container coords
                GetCursorPos cTopLeftPos
                ClientToScreen cTmrHwnd, mPoint
                cTopLeftPos.X = cTopLeftPos.X - mPoint.X - X
                cTopLeftPos.y = cTopLeftPos.y - mPoint.y - y
                ' trigger a mouse enter event
                SetProp cProjOwner, cPropKey & "Capture", myPointer
                RaiseEvent MouseEnter
                SetTimer cTmrHwnd, myPointer, 100, cTmrAddrOf
                UserControl.Tag = "Timer"
            End If
        Case Else   ' some other control has the mouse, but no longer cause we do
            
            Dim cAIC As aicAlphaImage
            Dim objAIC As aicAlphaImage
            ' get soft copy of the control
            CopyMemory cAIC, hCapturedAIC, 4&
            ' convert it to hard copy in case user has END statement or something just as bad
            ' in the mouse exit event. Then destroy softcopy
            Set objAIC = cAIC
            CopyMemory cAIC, 0&, 4&
            ' call a mouse exit for the other control
            Call objAIC.iccRemoteMouseExit
            Set objAIC = Nothing
            
            ' now update the window property and call mouse enter if needed
            If HitResult = vbHitResultHit Then
                ' convert screen coordinates to relative container coords
                GetCursorPos cTopLeftPos
                ClientToScreen cTmrHwnd, mPoint
                cTopLeftPos.X = cTopLeftPos.X - mPoint.X - X
                cTopLeftPos.y = cTopLeftPos.y - mPoint.y - y
                ' firre the mouse enter event
                SetProp cProjOwner, cPropKey & "Capture", myPointer
                RaiseEvent MouseEnter
                SetTimer cTmrHwnd, myPointer, 100, cTmrAddrOf
                UserControl.Tag = "Timer"
            Else
                SetProp cProjOwner, cPropKey & "Capture", 0&
            End If
        End Select
    End If
End Sub

Private Sub UserControl_Paint()
    
    
    If cScaledCoords.Width < 0 Then Exit Sub
    
    Dim uRect As RECT, iRect As RECT, clipRect As RECT
    
    GetClipBox UserControl.hdc, uRect
    
    If cRotated Then
        SetRect iRect, cScaledCoords.left, cScaledCoords.tOp, cScaledCoords.left + cScaledCoords.RotatedSize, cScaledCoords.tOp + cScaledCoords.RotatedSize
    Else
        SetRect iRect, cScaledCoords.left, cScaledCoords.tOp, cScaledCoords.left + cScaledCoords.Width, cScaledCoords.tOp + cScaledCoords.Height
    End If
    If Not IntersectRect(clipRect, uRect, iRect) = 0 Then
    
        If Not cOffscreen Is Nothing Then       ' AutoRedraw=True, we have a sized copy, use it
            clipRect.Right = clipRect.Right - clipRect.left
            clipRect.Bottom = clipRect.Bottom - clipRect.tOp
        
            cOffscreen.Render UserControl.hdc, clipRect.left, clipRect.tOp, clipRect.Right, clipRect.Bottom, _
                clipRect.left, clipRect.tOp, clipRect.Right, clipRect.Bottom, cOpacity, , , , cGrayScale
            
        ElseIf cRotated Then  ' rotating; slowest rendering method
            cImage.RotateAtCenterPoint UserControl.hdc, cAngle, UserControl.ScaleWidth \ 2, UserControl.ScaleHeight \ 2, cScaledCoords.Width, cScaledCoords.Height, , , , , cOpacity, , cGrayScale
        
        ElseIf cScaledCoords.OneToOne Then ' fast rendering
            clipRect.Right = clipRect.Right - clipRect.left
            clipRect.Bottom = clipRect.Bottom - clipRect.tOp
        
            cImage.Render UserControl.hdc, clipRect.left, clipRect.tOp, clipRect.Right, clipRect.Bottom, _
                clipRect.left - cScaledCoords.left, clipRect.tOp - cScaledCoords.tOp, clipRect.Right, clipRect.Bottom, cOpacity, , , , cGrayScale
        
        Else    ' resized; besides rotation; next slowest rendering method
            cImage.Render UserControl.hdc, cScaledCoords.left, cScaledCoords.tOp, cScaledCoords.Width, cScaledCoords.Height, , , , , cOpacity, , , , cGrayScale
            
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    cOpacity = 100&
    cGrayScale = aiNoGrayScale - 1&
    cScaleMethod = aiScaled
    cMask.Color = vbButtonFace
    cKeyProps = 1 Or 4   ' set high quality to true & AutoSize to True by default
    cImage.HighQualityInterpolation = True
    On Error Resume Next
    ' attempt to get actual scalemode. Errors occur when container has no scalemode (i.e.,frame)
    cScaleMode = UserControl.Parent.ScaleMode
    If Err Then
        Err.Clear
        cScaleMode = vbContainerPosition
        '^^ the above can be used for any scalemode; however, when container scalemode is
        ' vbTwips or vbPixels, the returned X,Y coordinates passed back are decimal not integer
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim srcData() As Byte
    With PropBag
        srcData = .ReadProperty("Image", srcData)
        cScaleMethod = .ReadProperty("Scaler", aiScaled)
        cOpacity = .ReadProperty("Opacity", 100)
        cKeepOrigFormat = .ReadProperty("KeepFormat", False)
        cMirror = .ReadProperty("Mirror", aiMirrorNone)
        cAngle = .ReadProperty("Angle", 0&)
        cGrayScale = .ReadProperty("GrayScale", aiNoGrayScale - 1)
        cHitTest = .ReadProperty("HitTest", aiEntireControl)
        cKeyProps = .ReadProperty("Props", 1)
        cMask.Applied = .ReadProperty("MaskUsed", False)
        cMask.Color = .ReadProperty("MaskColor", vbButtonFace)
        cMask.Source = .ReadProperty("MaskSource", aiNoMask)
        If cMask.Applied Then cMask.AppliedColor = .ReadProperty("Mask", vbButtonFace)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set UserControl.MouseIcon = .ReadProperty("MIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MPointer", vbDefault)
        UserControl.OLEDropMode = .ReadProperty("OLEdrop", vbOLEDropNone)
        cRotated = (Not cAngle = 0) Or ((cAngle = 0 And (cKeyProps And 4) = 4))
        If cScaleMethod = aiLockScale Then
            cScaledCoords.Width = .ReadProperty("ScaleCx", 0&)
            cScaledCoords.Height = .ReadProperty("ScaleCy", 0&)
        End If
        
    End With
    
    On Error Resume Next
    ' attempt to get actual scalemode. Errors occur when container has no scalemode (i.e.,frame)
    cScaleMode = UserControl.Parent.ScaleMode
    If Err Then
        Err.Clear
        cScaleMode = vbContainerPosition
        '^^ the above can be used for any scalemode; however, when container scalemode is
        ' vbTwips or vbPixels, the returned X,Y coordinates passed back are decimal not integer
    End If
    
    ' load/validate the image, set it up iniitally or fall back to checkboard state
    cImage.HighQualityInterpolation = ((cKeyProps And 1) = 1)
    If cImage.LoadPicture_Stream(srcData, UserControl.ScaleWidth, UserControl.ScaleHeight, , , (Ambient.UserMode = False Or cKeepOrigFormat)) = True Then
        sptMirrorImage 0&
        If cMask.Applied Then cImage.MakeTransparent sptConvertSysColor(cMask.AppliedColor)
        If (cKeyProps And 64) = 64 Then cImage.MakeImageInverse
        If ((cKeyProps And 8) = 8) Then
            If Ambient.UserMode = True Then
                cKeyProps = cKeyProps Or 32
                If sptUpdateOffscreen(True, True) = False Then sptRefreshRegion
            End If
        End If
    Else
        If cScaleMethod = aiLockScale Then cScaleMethod = aiScaled
        cKeyProps = (cKeyProps And 1)
        cAngle = 0
        cOpacity = 100
        cMirror = aiMirrorNone
        cScaledCoords.Height = 0: cScaledCoords.Width = 0
    End If
    Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim srcData() As Byte
    If Not cImage.ImageType = imgCheckerBoard Then
        Select Case cImage.ImageType
            Case imgBitmap, imgBmpARGB, imgBmpPARGB
                If cImage.isGDIplusEnabled Or cImage.isZlibEnabled Then
                    sptMirrorImage 0&
                    If cMask.Applied = True Then sptUndoMask
                    If (cKeyProps And 64) = 64 Then cImage.MakeImageInverse
                    cImage.SaveToStream_PNG srcData
                End If
            Case Else
                cImage.GetOrginalFormat srcData
        End Select
    End If
        
    With PropBag
        .WriteProperty "Image", srcData
        .WriteProperty "Scaler", cScaleMethod, aiScaled
        .WriteProperty "Opacity", cOpacity, 100
        .WriteProperty "KeepFormat", cKeepOrigFormat, False
        .WriteProperty "Mirror", cMirror, aiMirrorNone
        .WriteProperty "Angle", cAngle, 0&
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "MPointer", UserControl.MousePointer, vbDefault
        .WriteProperty "OLEdrop", UserControl.OLEDropMode, vbOLEDropNone
        .WriteProperty "MIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "GrayScale", cGrayScale, aiNoGrayScale - 1
        .WriteProperty "HitTest", cHitTest, aiEntireControl
        .WriteProperty "Props", cKeyProps, 1
        .WriteProperty "MaskUsed", cMask.Applied, False
        .WriteProperty "MaskColor", cMask.Color, vbButtonFace
        .WriteProperty "MaskSource", cMask.Source, aiNoMask
        If cMask.Applied Then .WriteProperty "Mask", cMask.AppliedColor, vbButtonFace
        
        If cScaleMethod = aiLockScale Then
            .WriteProperty "ScaleCx", cScaledCoords.Width, 0&
            .WriteProperty "ScaleCy", cScaledCoords.Height, 0&
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    
    If cScaledCoords.Width < 0 Then Exit Sub
    
    If cImage.Handle = 0& Then
        If Ambient.UserMode = False Then
            cImage.InitializeDIB UserControl.ScaleWidth, UserControl.ScaleHeight
            cImage.CreateCheckerBoard
        End If
    End If
    
    GetImageScales cScaledCoords.Width, cScaledCoords.Height, cScaleMethod, _
                    UserControl.ScaleWidth, UserControl.ScaleHeight, cRotated
    
    With cScaledCoords
        If cRotated = True Then
            .RotatedSize = Sqr(.Width * .Width + .Height * .Height)
            If cScaleMethod = aiLockScale Then
                .left = (UserControl.ScaleWidth - .Width) \ 2
                If .left < 0 Then .left = 0
                .tOp = (UserControl.ScaleHeight - .Height) \ 2
                If .tOp < 0 Then .tOp = 0
            Else
                .left = (UserControl.ScaleWidth - .RotatedSize) \ 2
                .tOp = (UserControl.ScaleHeight - .RotatedSize) \ 2
            End If
            .OneToOne = False
        Else
            .left = (UserControl.ScaleWidth - .Width) \ 2
            If .left < 0 Then .left = 0
            .tOp = (UserControl.ScaleHeight - .Height) \ 2
            If .tOp < 0 Then .tOp = 0
            If .Height = cImage.Height Then
                .OneToOne = (.Width = cImage.Width)
            Else
                .OneToOne = False
            End If
        End If
    End With
    
    If ((cKeyProps And 4) = 4) Then
        Dim sizeCx As Long, sizeCy As Long
        If cRotated = True Then
            sizeCx = cScaledCoords.RotatedSize
            sizeCy = sizeCx
        Else
            sizeCx = cScaledCoords.Width
            sizeCy = cScaledCoords.Height
        End If
        cScaledCoords.left = 0: cScaledCoords.tOp = 0
        cScaledCoords.Width = -cScaledCoords.Width
        UserControl.Size ScaleX(sizeCx, vbPixels, vbTwips), ScaleY(sizeCy, vbPixels, vbTwips)
        cScaledCoords.Width = -cScaledCoords.Width
    End If
'    If cImage.Handle = 0& Then Stop
    If sptUpdateOffscreen(True, True) = False Then sptRefreshRegion
    
End Sub


'============================================================================================
' /////////////////// C A L L B A C K   T H U N K I N G   R O U T I N E S \\\\\\\\\\\\\\\\\\\
'============================================================================================

'*************************************************************************************************
'* cCallback - Class generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

'-Callback code-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  zb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  j = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < j
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    If Not z_CbMem = 0 Then                                            'If memory allocated
        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
            z_CbMem = 0  'Release; Indicate memory released
            Erase z_Cb()
        End If
    End If
End Sub

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 2
Private Function Timer_Fader(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long

    KillTimer hwnd, TimerID    ' stop current timer
    
    ' ensure new opacity does not exceed final opacity
    If cFader.fStep < 0 Then    ' are we fading out? else we are fading in
        If cFader.fStep + cOpacity <= cFader.fOpacity Then cFader.fStep = 0&
    Else
        If cFader.fStep + cOpacity >= cFader.fOpacity Then cFader.fStep = 0&
    End If
    
    If cFader.fStep = 0& Then   ' fade to next step
        Me.Opacity = cFader.fOpacity
        RaiseEvent FadeTerminated(cOpacity)
    Else
        Me.Opacity = cOpacity + cFader.fStep
        SetTimer hwnd, TimerID, cFader.fDelay, cFader.tmrAddr
    End If
    
End Function

'Callback ordinal 1
Private Function Timer_MouseExit(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    KillTimer hwnd, TimerID    ' stop current timer
    
    Dim mHwnd As Long
    Dim tPoint As POINTAPI, mPoint As POINTAPI
    Dim bReset As Boolean
    
    ' validate that our control still thinks it has the mouse
    If GetProp(cProjOwner, cPropKey & "Capture") = ObjPtr(Me) Then
    
        GetCursorPos mPoint                             ' get current mouse position (screen coords)
        mHwnd = WindowFromPoint(mPoint.X, mPoint.y)     ' see if mouse is over control's container
        If mHwnd = hwnd Then
            ClientToScreen hwnd, tPoint
            ' adjust the points to control coordinates vs screen coordinates
            tPoint.X = mPoint.X - tPoint.X - cTopLeftPos.X
            tPoint.y = mPoint.y - tPoint.y - cTopLeftPos.y
            
            If cRegion = 0& Then
                ' we don't have a region, so the entire control is valid for hit testing
                If Not PtInRect(cRgnBox, tPoint.X, tPoint.y) Then bReset = True   ' restart the timer, no change
            Else
                ' we do have a region, see if the point is still within that region
                If Not PtInRegion(cRegion, tPoint.X, tPoint.y) = 0 Then bReset = True
            End If
        End If
        If bReset Then
            ' set timer for next check
            SetTimer hwnd, TimerID, 100, cTmrAddrOf
        Else
            ' mouse no longer over the control's region/shape; fire MouseExit
            SetProp cProjOwner, cPropKey & "Capture", 0&
            UserControl.Tag = vbNullString
            RaiseEvent MouseExit
        End If
    Else    ' control lost mouse before timer fired; do nothing
        UserControl.Tag = vbNullString
    End If
    
eh:
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the class timer's will fail every time.
End Function
