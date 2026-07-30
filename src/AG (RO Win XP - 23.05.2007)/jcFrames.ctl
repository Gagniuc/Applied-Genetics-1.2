VERSION 5.00
Begin VB.UserControl jcFrames 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   HitBehavior     =   0  'None
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "jcFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Original Code by Juan Carlos San Román - Planet Source Code

Option Explicit

'*************************************************************
'   Required Type Definitions
'*************************************************************

'Mover o form
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

'Mudar a borda para mudar tamanho
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Private Type Point
    X As Long
    y As Long
End Type

Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type

Public Enum jcStyleConst
    XPDefault = 0
    jcGradient = 1
    TextBox = 2
    Windows = 3
    Messenger = 4
    InnerWedge = 5
    OuterWedge = 6
    Header = 7
    Panel = 8
End Enum

'xp theme
Public Enum jcThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    Custom = 5
    xThemeDarkBlue = 6
    xThemeGreen = 7
    xThemeOffice2003Style2 = 8
    xThemeMetallic = 9
    xThemeOrange = 10
    xThemeTurquoise = 11
    xThemeGray = 12
    xThemeDarkBlue2 = 13
    xThemeMoney = 14
    xThemeOffice2003Style1 = 15
End Enum

'Responsável por mover o form
Public Enum jcResp
    jcTitulo = 0
    jcPainel = 1
    jcAmbos = 2
End Enum

'gradient type
Public Enum jcGradConst
    VerticalGradient = 0
    horizontalGradient = 1
    VCilinderGradient = 2
    hCilinderGradient = 3
End Enum

'header style
Public Enum jcHeaderConst
    TxtBoxColor = 0
    Gradient = 1
End Enum

'TxtBox style
Public Enum jcShadowConst
    [No shadow] = 0
    Shadow = 1
End Enum

'gradient backstyle
Public Enum jcBackStyleConst
    g_Opaque = 0
    g_Transparent = 1
End Enum

'icon aligment
Public Enum IconAlignConst
    vbLeftAligment = 0
    vbRightAligment = 1
End Enum

Enum m_PanelArea
    xTitle = 0
    xPanel = 1
End Enum

'*************************************************************
'   events
'*************************************************************
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Event TileClick()
Event PanelClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single, PanelArea As m_PanelArea)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single, PanelArea As m_PanelArea)
Event CollapseClick(Button As Integer)

'*************************************************************
'   Required API Declarations
'*************************************************************

'FORM TRANSPARENTE
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

'VERSÃO DO WINDOWS
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As Point) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal left As Long, ByVal tOp As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private useMask As Boolean, useGrey As Boolean

'*************************************************************
'   Members
'*************************************************************
Private m_FrameColor            As OLE_COLOR
Private m_TextBoxColor          As OLE_COLOR
Private m_BackColor             As OLE_COLOR
Private m_FillColor             As OLE_COLOR

Private m_FrameColorDis         As OLE_COLOR
Private m_TextBoxColorDis       As OLE_COLOR
Private m_FillColorDis          As OLE_COLOR
Private jcColorToDis            As OLE_COLOR
Private jcColorFromDis          As OLE_COLOR
Private jcColorBorderPicDis     As OLE_COLOR

Private m_FrameColorIni         As OLE_COLOR
Private m_TextBoxColorIni       As OLE_COLOR
Private m_FillColorIni          As OLE_COLOR
Private jcColorToIni            As OLE_COLOR
Private jcColorFromIni          As OLE_COLOR
Private jcColorBorderPicIni     As OLE_COLOR

Private m_Caption               As String
Private m_Enabled               As Boolean
Private m_TextBoxHeight         As Long
Private m_TextHeight            As Long
Private m_TextWidth             As Long
Private m_Height                As Long
Private m_TextColor             As Long
Private m_Alignment             As Long
Private m_Font                  As StdFont
Private m_RoundedCorner         As Boolean
Private m_RoundedCornerTxtBox   As Boolean
Private m_style                 As jcStyleConst
Private m_HeaderStyle           As jcHeaderConst
Private m_GradientHeaderStyle   As jcGradConst
Private m_icon                  As StdPicture
Private m_IconSize              As Integer
Private m_IconAlignment         As IconAlignConst
Private m_ThemeColor            As jcThemeConst
Private m_ColorTo               As OLE_COLOR
Private m_ColorFrom             As OLE_COLOR
Private m_Indentation           As Integer
Private m_Space                 As Integer
Private m_TxtBoxShadow          As jcShadowConst
Private jcTextBoxCenter         As Long
Private jcTextDrawParams        As Long
Private jcColorTo               As OLE_COLOR
Private jcColorFrom             As OLE_COLOR
Private jcColorBorderPic        As OLE_COLOR
Private jcLpp As Point
Private Const TEXT_INACTIVE = &H80000011 '&H6A6A6A
Private Const m_Border_Inactive = &H8000000B
Private Const m_BtnFace_Inactive = &H8000000F
Private Const m_BtnFace = &H80000016 '&H8000000F '&H80000016&

'*************************************************************
'   Constants
'*************************************************************
Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100
Private Const DT_CALCRECT = &H400

Private Const ALTERNATE = 1      ' ALTERNATE and WINDING are
Private Const WINDING = 2        ' constants for FillMode.
Private Const BLACKBRUSH = 4     ' Constant for brush type.
Private Const WHITE_BRUSH = 0    ' Constant for brush type.

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4

Private Const m_def_Responsavel = 0
Private Const m_def_AllowDraging = 0
Private Const m_def_AtivarResizeDoForm = False
Private Const m_def_Collapsar = False

Dim temp_height As Integer
Public Collapsado As Boolean

Dim m_Responsavel As jcResp
Dim m_AllowDraging As Boolean
Dim m_AllowParentDraging As Boolean
Dim m_AtivarResizeDoForm As Boolean

Dim m_Collapsar As Boolean

'Eventos
Function PanelOrTitle(y As Single) As m_PanelArea
    If (y <= 0) Or (y < m_TextBoxHeight) Then
        PanelOrTitle = xTitle
    Else
        PanelOrTitle = xPanel
    End If
End Function


Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent CollapseClick(Button)
    If (Button = 1) Then
        Dim iY As Integer
        If m_Collapsar Then
            Select Case m_style
                Case Messenger: iY = 9
                Case jcGradient: iY = 8
                Case XPDefault: iY = 6
                Case InnerWedge: iY = -1
                Case OuterWedge: iY = -2
                Case Windows: iY = 1
            End Select
            If Collapsado Then
                Label.Caption = "6"
                Collapsado = False
                UserControl.Height = temp_height
            Else
                Label.Caption = "5"
                Collapsado = True
                temp_height = UserControl.Height
                UserControl.Height = ScaleY(m_TextBoxHeight + iY, vbPixels, vbTwips)
            End If
        End If
        PaintFrame
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, X, y, PanelOrTitle(y))
    If (Button <> vbLeftButton) Then Exit Sub
    Select Case PanelOrTitle(y)
        Case xTitle
            RaiseEvent TileClick
        Case xPanel
            RaiseEvent PanelClick
    End Select
    
    Dim iHwnd As Long
    If m_AllowDraging Then
        If m_AllowParentDraging Then
            If (m_Responsavel = jcAmbos) Or (m_Responsavel = PanelOrTitle(y)) Then
                iHwnd = UserControl.Parent.hwnd
                jcTransp 70
            End If
        Else
            iHwnd = UserControl.hwnd
        End If
        Call ReleaseCapture
        Call SendMessage(iHwnd, &HA1, 2, 0&)
        If m_AllowParentDraging Then jcTransp 255
    End If
    
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, X, y, PanelOrTitle(y))
End Sub


'==========================================================================
' Init, Initialize, Read & Write UserControl
'==========================================================================
Private Sub UserControl_InitProperties()
    'Set default properties
    m_Caption = Ambient.DisplayName
    m_Responsavel = m_def_Responsavel
    m_AllowDraging = m_def_AllowDraging
    m_AtivarResizeDoForm = m_def_AtivarResizeDoForm
    m_AllowParentDraging = False
    m_Collapsar = m_def_Collapsar
    m_BackColor = TranslateColor(Ambient.BackColor)
    m_FillColorIni = TranslateColor(Ambient.BackColor)
    m_RoundedCorner = True
    m_RoundedCornerTxtBox = False
    m_style = jcGradient
    m_ThemeColor = Blue
    m_TextColor = TranslateColor(vbBlack)
    m_FrameColorIni = TranslateColor(vbBlack)
    m_TextBoxColorIni = TranslateColor(vbWhite)
    m_TxtBoxShadow = [No shadow]
    m_TextBoxHeight = 22
    m_HeaderStyle = Gradient
    m_GradientHeaderStyle = VerticalGradient
    SetjcTextDrawParams
End Sub

Private Sub UserControl_Initialize()
    Set m_Font = New StdFont
    Set UserControl.Font = m_Font
    m_IconSize = 16
    m_ColorFrom = 10395391
    m_ColorTo = 15790335
    m_TxtBoxShadow = [No shadow]
    m_ThemeColor = Blue
    m_Enabled = True
    Call SetDefaultThemeColor(m_ThemeColor)
    m_TextBoxHeight = 22
    m_Alignment = vbCenter
    m_IconAlignment = vbLeftAligment
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'On Error Resume Next
    With PropBag
        m_FrameColorIni = .ReadProperty("FrameColor", vbBlack)
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_AtivarResizeDoForm = PropBag.ReadProperty("AtivarResizeDoForm", m_def_AtivarResizeDoForm)
        m_Responsavel = PropBag.ReadProperty("MoverResponsavel", m_def_Responsavel)
        m_AllowDraging = PropBag.ReadProperty("MoverControle", m_def_AllowDraging)
        m_Collapsar = PropBag.ReadProperty("Collapsar", m_def_Collapsar)
        m_AllowParentDraging = PropBag.ReadProperty("MoverForm", False)
        m_FillColorIni = .ReadProperty("FillColor", Ambient.BackColor)
        m_TextBoxColorIni = .ReadProperty("TextBoxColor", vbWhite)
        m_TxtBoxShadow = .ReadProperty("TxtBoxShadow", [No shadow])
        m_style = .ReadProperty("Style", jcGradient)
        m_RoundedCorner = .ReadProperty("RoundedCorner", True)
        m_Enabled = .ReadProperty("Enabled", True)
        m_RoundedCornerTxtBox = .ReadProperty("RoundedCornerTxtBox", False)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_TextBoxHeight = .ReadProperty("TextBoxHeight", 22)
        m_TextColor = .ReadProperty("TextColor", vbBlack)
        m_Alignment = .ReadProperty("Alignment", vbCenter)
        m_IconAlignment = .ReadProperty("IconAlignment", vbLeftAligment)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        Set m_icon = .ReadProperty("Picture", Nothing)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_ThemeColor = .ReadProperty("ThemeColor", Blue)
        m_ColorFrom = .ReadProperty("ColorFrom", 10395391)
        m_ColorTo = .ReadProperty("ColorTo", 15790335)
        m_HeaderStyle = .ReadProperty("HeaderStyle", TxtBoxColor)
        m_GradientHeaderStyle = .ReadProperty("GradientHeaderStyle", VerticalGradient)
    End With
    
    If m_Collapsar Then
        Select Case m_style
            Case Header, Panel, TextBox: Label.Visible = False
            Case Else: Label.Visible = True
        End Select
    End If
    
    'Add properties
    UserControl.BackColor = TranslateColor(m_BackColor)
    SetjcTextDrawParams
    SetFont m_Font
    Call SetDefaultThemeColor(m_ThemeColor)
    SetDisabledColor
    'Paint control
    PaintFrame
    
    If m_AtivarResizeDoForm Then
        If Ambient.UserMode Then
            Set frm = Parent
            SetWindowLong frm.hwnd, GWL_STYLE, GetWindowLong(frm.hwnd, GWL_STYLE) And Not (WS_CAPTION)
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FrameColor", m_FrameColorIni, vbBlack
        .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
        .WriteProperty "FillColor", m_FillColorIni, Ambient.BackColor
        .WriteProperty "TextBoxColor", m_TextBoxColorIni, vbWhite
        .WriteProperty "MoverForm", m_AllowParentDraging, False
        .WriteProperty "MoverControle", m_AllowDraging, m_def_AllowDraging
        .WriteProperty "AtivarResizeDoForm", m_AtivarResizeDoForm, m_def_AtivarResizeDoForm
        .WriteProperty "MoverResponsavel", m_Responsavel, m_def_Responsavel
        .WriteProperty "Collapsar", m_Collapsar, m_def_Collapsar
        .WriteProperty "TxtBoxShadow", m_TxtBoxShadow, [No shadow]
        .WriteProperty "Style", m_style, jcGradient
        .WriteProperty "RoundedCorner", m_RoundedCorner, True
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "RoundedCornerTxtBox", m_RoundedCornerTxtBox, False
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "TextBoxHeight", m_TextBoxHeight, 22
        .WriteProperty "TextColor", m_TextColor, vbBlack
        .WriteProperty "Alignment", m_Alignment, vbCenter
        .WriteProperty "IconAlignment", m_IconAlignment, vbLeftAligment
        .WriteProperty "Font", m_Font, Ambient.Font
        .WriteProperty "Picture", m_icon, Nothing
        .WriteProperty "IconSize", m_IconSize, 16
        .WriteProperty "ThemeColor", m_ThemeColor, Blue
        .WriteProperty "ColorFrom", m_ColorFrom, 10395391
        .WriteProperty "ColorTo", m_ColorTo, 15790335
        .WriteProperty "HeaderStyle", m_HeaderStyle, TxtBoxColor
        .WriteProperty "GradientHeaderStyle", m_GradientHeaderStyle, VerticalGradient
    End With
End Sub

'==========================================================================
' Usercontrol events
'==========================================================================

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub UserControl_Resize()
    If (m_Collapsar = False) Then
        If UserControl.Width < 700 Then UserControl.Width = 700
        If UserControl.Height < 400 Then UserControl.Height = 400
    End If
    PaintFrame
End Sub

'==========================================================================
' Properties
'==========================================================================
Public Property Let FrameColor(ByRef new_FrameColor As OLE_COLOR)
    m_FrameColorIni = TranslateColor(new_FrameColor)
    If m_ThemeColor = Custom Then jcColorBorderPic = m_FrameColor
    PropertyChanged "FrameColor"
    PaintFrame
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = m_FrameColorIni
End Property

Public Property Let FillColor(ByRef new_FillColor As OLE_COLOR)
Attribute FillColor.VB_Description = "Returns/Sets the Fill color for TextBox and Windows style"
    m_FillColorIni = TranslateColor(new_FillColor)
    PropertyChanged "FillColor"
    PaintFrame
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColorIni
End Property

Public Property Let RoundedCornerTxtBox(ByRef new_RoundedCornerTxtBox As Boolean)
    m_RoundedCornerTxtBox = new_RoundedCornerTxtBox
    PropertyChanged "RoundedCornerTxtBox"
    PaintFrame
End Property

Public Property Get RoundedCornerTxtBox() As Boolean
    RoundedCornerTxtBox = m_RoundedCornerTxtBox
End Property

Public Property Let Enabled(ByRef New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    PaintFrame
    FrameEnabled m_Enabled
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let TxtBoxShadow(ByRef new_TxtBoxShadow As jcShadowConst)
    m_TxtBoxShadow = new_TxtBoxShadow
    PropertyChanged "TxtBoxShadow"
    PaintFrame
End Property

Public Property Get TxtBoxShadow() As jcShadowConst
    TxtBoxShadow = m_TxtBoxShadow
End Property

Public Property Let RoundedCorner(ByRef new_RoundedCorner As Boolean)
    m_RoundedCorner = new_RoundedCorner
    PropertyChanged "RoundedCorner"
    PaintFrame
End Property

Public Property Get RoundedCorner() As Boolean
    RoundedCorner = m_RoundedCorner
End Property

Public Property Let Caption(ByRef New_Caption As String)
    m_Caption = New_Caption
    PaintFrame
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Alignment(ByRef New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    SetjcTextDrawParams
    PropertyChanged "Alignment"
    PaintFrame
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Style(ByRef New_Style As jcStyleConst)
    m_style = New_Style
    PropertyChanged "Style"
    If (New_Style = Header) Or (New_Style = Panel) Or (New_Style = TextBox) Then m_Collapsar = False: Label.Visible = False
    SetDefault ' m_ThemeColor
    PaintFrame
End Property

Public Property Get Style() As jcStyleConst
    If (m_style = Header) Or (m_style = Panel) Or (m_style = TextBox) Then m_Collapsar = False: Label.Visible = False
    Style = m_style
End Property

Public Property Let TextBoxHeight(ByRef new_TextBoxHeight As Long)
    m_TextBoxHeight = new_TextBoxHeight
    PropertyChanged "TextBoxHeight"
    PaintFrame
End Property

Public Property Get TextBoxHeight() As Long
    TextBoxHeight = m_TextBoxHeight
End Property

Public Property Let TextColor(ByRef new_TextColor As OLE_COLOR)
    m_TextColor = TranslateColor(new_TextColor)
    PropertyChanged "TextColor"
    PaintFrame
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextBoxColor(ByRef new_TextBoxColor As OLE_COLOR)
    m_TextBoxColorIni = TranslateColor(new_TextBoxColor)
    PropertyChanged "TextBoxColor"
    PaintFrame
End Property

Public Property Get TextBoxColor() As OLE_COLOR
    TextBoxColor = m_TextBoxColorIni
End Property

Public Property Let BackColor(ByRef New_BackColor As OLE_COLOR)
    m_BackColor = TranslateColor(New_BackColor)
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
    PaintFrame
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Set Font(ByRef New_Font As StdFont)
    SetFont New_Font
    PropertyChanged "Font"
    PaintFrame
End Property

Public Property Let Font(ByRef New_Font As StdFont)
    SetFont New_Font
    PropertyChanged "Font"
    PaintFrame
End Property
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_icon
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_icon = New_Picture
    PropertyChanged "Picture"
    PaintFrame
End Property

Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Value As Integer)
    m_IconSize = New_Value
    PropertyChanged "IconSize"
    PaintFrame
End Property

Public Property Let IconAlignment(ByRef new_IconAlignment As IconAlignConst)
    m_IconAlignment = new_IconAlignment
    PropertyChanged "IconAlignment"
    PaintFrame
End Property

Public Property Get IconAlignment() As IconAlignConst
    IconAlignment = m_IconAlignment
End Property

Public Property Get ThemeColor() As jcThemeConst
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As jcThemeConst)
    If m_ThemeColor <> vData Then
        m_ThemeColor = vData
        Call SetDefaultThemeColor(m_ThemeColor)
        PaintFrame
        PropertyChanged "ThemeColor"
    End If
End Property

Public Property Get ColorFrom() As OLE_COLOR
Attribute ColorFrom.VB_Description = "Returns/Sets the Start color for gradient"
    ColorFrom = m_ColorFrom
End Property

Public Property Let ColorFrom(ByRef new_ColorFrom As OLE_COLOR)
    m_ColorFrom = TranslateColor(new_ColorFrom)
    If m_ThemeColor = Custom Then jcColorFromIni = m_ColorFrom
    PropertyChanged "ColorFrom"
    PaintFrame
End Property

Public Property Get ColorTo() As OLE_COLOR
Attribute ColorTo.VB_Description = "Returns/Sets the End color for gradient"
    ColorTo = m_ColorTo
End Property

Public Property Let ColorTo(ByRef new_ColorTo As OLE_COLOR)
    m_ColorTo = TranslateColor(new_ColorTo)
    If m_ThemeColor = Custom Then jcColorToIni = m_ColorTo
    PropertyChanged "ColorTo"
    PaintFrame
End Property

Public Property Let HeaderStyle(ByRef new_HeaderStyle As jcHeaderConst)
    m_HeaderStyle = new_HeaderStyle
    PropertyChanged "HeaderStyle"
    PaintFrame
End Property

Public Property Get AtivarResizeDoForm() As Boolean
    AtivarResizeDoForm = m_AtivarResizeDoForm
End Property

Public Property Let AtivarResizeDoForm(ByVal New_AtivarResizeDoForm As Boolean)
    m_AtivarResizeDoForm = New_AtivarResizeDoForm
    PropertyChanged "AtivarResizeDoForm"
End Property

Public Property Get MoverResponsavel() As jcResp
    MoverResponsavel = m_Responsavel
End Property

Public Property Let MoverResponsavel(ByVal New_Responsavel As jcResp)
    m_Responsavel = New_Responsavel
    PropertyChanged "MoverResponsavel"
End Property

Public Property Get MoverControle() As Boolean
    MoverControle = m_AllowDraging
End Property

Public Property Let MoverControle(ByVal New_AllowDraging As Boolean)
    m_AllowDraging = New_AllowDraging
    PropertyChanged "MoverControle"
End Property

Public Property Get MoverForm() As Boolean
    MoverForm = m_AllowParentDraging
End Property

Public Property Let MoverForm(ByVal New_Drag As Boolean)
    m_AllowParentDraging = New_Drag
    PropertyChanged "MoverForm"
    Call PaintFrame
End Property


'Faz o form aumentar e diminuir o seu tamanho (Collapsar)
Public Property Get Collapsar() As Boolean
    Collapsar = m_Collapsar
End Property
Public Property Let Collapsar(ByVal New_Drag As Boolean)
    Select Case m_style
        Case Header, Panel, TextBox:
            m_Collapsar = False
        Case Else:
            m_Collapsar = New_Drag
    End Select
    If m_Collapsar Then Label.Visible = True Else Label.Visible = False
    PropertyChanged "Collapsar"
End Property


Public Property Get HeaderStyle() As jcHeaderConst
    HeaderStyle = m_HeaderStyle
End Property

Public Property Let GradientHeaderStyle(ByRef new_GradientHeaderStyle As jcGradConst)
    m_GradientHeaderStyle = new_GradientHeaderStyle
    PropertyChanged "GradientHeaderStyle"
    PaintFrame
End Property

Public Property Get GradientHeaderStyle() As jcGradConst
    GradientHeaderStyle = m_GradientHeaderStyle
End Property

Private Sub SetjcTextDrawParams()
    'Set text draw params using m_Alignment
    If m_style = Panel Then
        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_WORDBREAK Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_WORDBREAK Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_WORDBREAK Or DT_VCENTER
        End If
    Else
        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
        End If
    End If
End Sub

Private Sub SetFont(ByRef New_Font As StdFont)
    With m_Font
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Name = New_Font.Name
        .Size = New_Font.Size
    End With
    Set UserControl.Font = m_Font
End Sub

'==========================================================================
' Functions and subroutines
'==========================================================================
Private Sub SetDefaultThemeColor(ThemeType As Long)
    Select Case ThemeType
        Case 0 '"Blue"
            jcColorFromIni = RGB(129, 169, 226)
            jcColorToIni = RGB(221, 236, 254)
            jcColorBorderPicIni = RGB(0, 0, 128)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(192, 212, 240)
                m_FillColorIni = RGB(192, 212, 240)
            End If
        Case 1 '"Silver"
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(204, 203, 218)
                m_FillColorIni = RGB(204, 203, 218)
            End If
        Case 2 '"Olive"
            jcColorFromIni = RGB(181, 197, 143)
            jcColorToIni = RGB(247, 249, 225)
            jcColorBorderPicIni = RGB(63, 93, 56)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(224, 224, 213)
                m_FillColorIni = RGB(224, 224, 213)
            End If
        Case 3 '"Visual2005"
            jcColorFromIni = RGB(194, 194, 171)
            jcColorToIni = RGB(248, 248, 242)
            jcColorBorderPicIni = RGB(145, 145, 115)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(224, 224, 213)
                m_FillColorIni = RGB(224, 224, 213)
            End If
        Case 4 '"Norton2004"
            jcColorFromIni = RGB(217, 172, 1)
            jcColorToIni = RGB(255, 239, 165)
            jcColorBorderPicIni = RGB(117, 91, 30)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(236, 214, 128)
                m_FillColorIni = RGB(236, 214, 128)
            End If
        Case 5  'Custom
            jcColorFromIni = m_ColorFrom
            jcColorToIni = m_ColorTo
            jcColorBorderPicIni = m_FrameColor
        Case 6
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(137, 170, 224)
            jcColorToIni = RGB(7, 33, 100)
            jcColorBorderPicIni = RGB(100, 144, 88)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(215, 230, 251)
                m_FillColorIni = RGB(142, 179, 231) '=
                m_BackColor = RGB(142, 179, 231) '=
            End If
        Case 7 'xThemeGreen
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(228, 235, 200)
            jcColorToIni = RGB(175, 194, 142)
            jcColorBorderPicIni = RGB(100, 144, 88)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(100, 144, 88)
                m_FillColorIni = RGB(233, 244, 207) '=
                m_BackColor = RGB(233, 244, 207) '=
            End If
        Case 8 'xThemeOffice2003Style2
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(249, 249, 255)
            jcColorToIni = RGB(159, 157, 185)
            jcColorBorderPicIni = RGB(124, 124, 148)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(110, 109, 143)
                m_FillColorIni = RGB(253, 250, 255) '=
                m_BackColor = RGB(253, 250, 255) '=
            End If
        Case 9 'xThemeMetallic
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(219, 220, 232)
            jcColorToIni = RGB(149, 147, 177)
            jcColorBorderPicIni = RGB(119, 118, 151)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(119, 118, 151)
                m_FillColorIni = RGB(232, 232, 232) '=
                m_BackColor = RGB(232, 232, 232) '=
            End If
        Case 10 'xThemeOrange
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(255, 122, 0)
            jcColorToIni = RGB(130, 0, 0)
            jcColorBorderPicIni = RGB(139, 0, 0)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(255, 222, 173)
                m_FillColorIni = RGB(255, 222, 173) '=
                m_BackColor = RGB(255, 222, 173) '=
            End If
        Case 11 'xThemeTurquoise
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(72, 209, 204)
            jcColorToIni = RGB(43, 103, 109)
            jcColorBorderPicIni = RGB(65, 131, 111)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(233, 250, 248)
                m_FillColorIni = RGB(224, 255, 255) '=
                m_BackColor = RGB(224, 255, 255) '=
            End If
        Case 12 'xThemeGray
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(192, 192, 192)
            jcColorToIni = RGB(51, 51, 51)
            jcColorBorderPicIni = RGB(51, 51, 51)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(235, 235, 235)
                m_FillColorIni = RGB(235, 235, 235) '=
                m_BackColor = RGB(235, 235, 235) '=
            End If
        Case 13 'xThemeDarkBlue2
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(81, 128, 208)
            jcColorToIni = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
            jcColorBorderPicIni = RGB(0, 45, 150)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = vbRed
                m_FillColorIni = RGB(142, 179, 231) '=
                m_BackColor = RGB(142, 179, 231) '=
            End If
        Case 14 'xThemeMoney
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(160, 160, 160)
            jcColorToIni = dBlendColor(RGB(90, 90, 90), vbBlack, 230)
            jcColorBorderPicIni = RGB(68, 68, 68)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = vbWhite
                m_FillColorIni = RGB(112, 112, 112) '=
                m_BackColor = RGB(112, 112, 112) '=
            End If
        Case 15 'xThemeOffice2003Style1
            m_GradientHeaderStyle = VerticalGradient: m_HeaderStyle = Gradient
            jcColorFromIni = RGB(209, 227, 251)
            jcColorToIni = RGB(106, 140, 203)
            jcColorBorderPicIni = RGB(0, 0, 128)
            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(110, 109, 143)
                m_FillColorIni = RGB(255, 255, 255) '=
                m_BackColor = RGB(255, 255, 255) '=
            End If
        Case Else
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)
    End Select
End Sub


Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR As Long

    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property


'==================
' Main drawing sub
'==================
Private Sub PaintFrame()
    Dim R As RECT, R_Caption As RECT, Rc As RECT
    Dim iX As Integer, iY As Integer
    
    m_Height = 3
    m_Indentation = 15
    m_Space = 6
    iX = 0
    iY = 0
    EraseRegion
    
    'Clear user control
    UserControl.Cls
    
    'Set caption height and width
    If Len(m_Caption) <> 0 Then
        m_TextWidth = UserControl.TextWidth(m_Caption)
        m_TextHeight = UserControl.TextHeight(m_Caption)
        jcTextBoxCenter = m_TextHeight / 2
    Else
        jcTextBoxCenter = 0
    End If
    
    'Select colors according to enabled property
    If m_Enabled = False Then
        m_FrameColor = m_FrameColorDis
        m_TextBoxColor = m_TextBoxColorDis
        m_FillColor = m_FillColorDis
        jcColorTo = jcColorToDis
        jcColorFrom = jcColorFromDis
        jcColorBorderPic = jcColorBorderPicDis
    Else
        m_FrameColor = m_FrameColorIni
        m_TextBoxColor = m_TextBoxColorIni
        m_FillColor = m_FillColorIni
        jcColorTo = jcColorToIni
        jcColorFrom = jcColorFromIni
        jcColorBorderPic = jcColorBorderPicIni
    End If
  
    'select frame style
    Select Case m_style
        Case Is = XPDefault
            Draw_XPDefault R_Caption
        Case Is = jcGradient
            Draw_jcGradient R_Caption, iY
        Case Is = TextBox
            Draw_TextBox R_Caption, iX, iY
        Case Is = Windows
            Draw_Windows R_Caption, iY
        Case Is = Messenger
            Draw_Messenger R_Caption, iY
        Case Is = InnerWedge
            Draw_InnerWedge R_Caption, iY
        Case Is = OuterWedge
            Draw_OuterWedge R_Caption, iY
        Case Is = Header
            Draw_Header R_Caption
        Case Is = Panel
            Draw_Panel R_Caption, iY
        Case Else
            Draw_jcGradient R_Caption, iY
    End Select
    
    'caption and icon alignments
    If Not (m_icon Is Nothing Or m_style = XPDefault) Then
        If m_IconAlignment = vbLeftAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.left = R_Caption.left + m_Space + m_IconSize
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.left = R_Caption.left + m_Space + m_IconSize
            Else
                R_Caption.left = R_Caption.left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            If m_style = TextBox Then
                iX = m_Indentation + m_Space * 2
            Else
                iX = m_Space
            End If
        ElseIf m_IconAlignment = vbRightAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            Else
                R_Caption.left = R_Caption.left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            If m_style = TextBox Then
                iX = UserControl.ScaleWidth - m_Space * 2 - m_IconSize - m_Indentation
            Else
                iX = UserControl.ScaleWidth - m_Space - m_IconSize
            End If
        End If
    End If

    'Draw caption
    If Len(m_Caption) <> 0 Then
        'Set text color
        Dim m_caption_aux As String
        m_caption_aux = TrimWord(m_Caption, R_Caption.Right - R_Caption.left)
        
        'Draw text
        UserControl.ForeColor = IIf(m_Enabled, m_TextColor, TranslateColor(TEXT_INACTIVE))
        
        If m_style = Panel Then
            CopyRect Rc, R_Caption
            DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), Rc, DT_CALCRECT Or DT_WORDBREAK, ByVal 0&
            OffsetRect Rc, (R_Caption.Right - Rc.Right) \ 2, (R_Caption.Bottom - Rc.Bottom) \ 2
            DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), Rc, jcTextDrawParams, ByVal 0&
        Else
            DrawTextEx UserControl.hdc, m_caption_aux, Len(m_caption_aux), R_Caption, jcTextDrawParams, ByVal 0&
        End If
        
    End If
    
    'draw picture
    If Not (m_icon Is Nothing Or m_style = XPDefault Or m_style = InnerWedge Or m_style = OuterWedge) Then
        If m_style = Messenger Then
            If iY < m_Height * 2 + 2 Then iY = m_Height * 2 + 2
        ElseIf m_style = jcGradient Then
            If iY < m_Height + 2 Then iY = m_Height + 2
        Else
            If iY < 0 Then iY = m_Space / 2
        End If
        
        If m_Enabled Then
            UserControl.PaintPicture m_icon, iX, iY, m_IconSize, m_IconSize
            'TransBlt UserControl.hDC, Ix, Iy, m_IconSize, m_IconSize, m_Icon, vbBlack, , , False, False
        Else
            TransBlt UserControl.hdc, iX, iY, m_IconSize, m_IconSize, m_icon, vbBlack, , , True, False
        End If
    End If
    
    Select Case m_style
        Case Messenger: iY = 5
        Case jcGradient: iY = 8
        Case XPDefault: iY = 6
        Case InnerWedge: iY = 11
        Case OuterWedge: iY = 11
        Case Windows: iY = 11
    End Select
    Label.Move UserControl.ScaleWidth - 30, CInt(ScaleY((ScaleY(m_TextBoxHeight, vbPixels, vbTwips) - Label.Height) / 2, vbTwips, vbPixels)) - iY
End Sub

Private Sub SetDefault()
    Select Case m_style
        Case Is = XPDefault
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = jcGradient
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            SetjcTextDrawParams
        Case Is = TextBox
            m_TextColor = vbBlack
            m_FrameColorIni = &H6A6A6A
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCornerTxtBox = True
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Windows
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCorner = True
            m_RoundedCornerTxtBox = False
            m_FillColorIni = &HE0FFFF
            m_GradientHeaderStyle = horizontalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams
        Case Is = Messenger
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams
        Case Is = InnerWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 192
            m_TextBoxColorIni = 192
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = OuterWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 10878976
            m_TextBoxColorIni = 10878976
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Header
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Panel
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VCilinderGradient
            m_HeaderStyle = Gradient
            SetjcTextDrawParams
    End Select
    
End Sub

Private Sub PaintShpInBar(iColorA As Long, iColorB As Long, m_Height As Long)
    Dim I As Integer, x_left As Integer, y_top As Integer, SpaceBtwnShp As Integer, NumShp As Integer
    Dim RectHeight As Long, RectWidth As Long, R As RECT

    SpaceBtwnShp = 2    'space between shapes
    NumShp = 9          'number of points
    RectHeight = 2      'shape height
    RectWidth = 2       'shape width
    
    'x and y shape  coordinates
    x_left = (UserControl.ScaleWidth - NumShp * RectWidth - (NumShp - 1) * SpaceBtwnShp) / 2
    y_top = (m_Height - RectHeight) / 2
    
    For I = 0 To NumShp - 1
        SetRect R, x_left + I * SpaceBtwnShp + I * RectWidth + 1, y_top + 1, 1, 1
        APIRectangle UserControl.hdc, R.left, R.tOp, R.Right, R.Bottom, iColorA
        SetRect R, x_left + I * SpaceBtwnShp + I * RectWidth, y_top, 1, 1
        APIRectangle UserControl.hdc, R.left, R.tOp, R.Right, R.Bottom, iColorB
    Next I
End Sub

Private Sub Draw_XPDefault(R_Caption As RECT)
    Dim p_left As Long, m_roundedRadius As Long, R As RECT, lpp As Point
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    If Len(m_Caption) <> 0 Then
        If m_Alignment = vbLeftJustify Then
            p_left = m_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Indentation - m_Space - 1
        Else
            p_left = (UserControl.ScaleWidth - 1 - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hdc, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor

        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If
End Sub

Private Sub Draw_jcGradient(R_Caption As RECT, iY As Integer)
    Dim R As RECT, m_roundedRadius As Long
    
    jcTextBoxCenter = m_TextBoxHeight / 2
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R

    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height
    DrawGradientInRectangle UserControl.hdc, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic
    
    If m_HeaderStyle = Gradient Then
         SetRect R, 0, m_Height, UserControl.ScaleWidth - 2, m_TextBoxHeight
         DrawGradientInRectangle UserControl.hdc, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
         SetRect R, 0, m_Height, UserControl.ScaleWidth - 1, m_TextBoxHeight + m_Height + 2
         DrawAPIRoundRect False, 0&, m_FillColor, m_FrameColor, R
    End If

    SetRect R, 0, m_Height + m_TextBoxHeight, UserControl.ScaleWidth - 2, m_Height
    DrawGradientInRectangle UserControl.hdc, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic

    SetRect R, 1, m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
    DrawGradientInRectangle UserControl.hdc, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
    
    'set caption rect
    SetRect R_Caption, m_Space, m_Height + 1, UserControl.ScaleWidth - 2 - m_Space, m_TextBoxHeight + 2

    'set icon Y coordinate
    iY = (m_Height * 2 + m_TextBoxHeight - m_IconSize) / 2
End Sub

Private Sub Draw_TextBox(R_Caption As RECT, iX As Integer, iY As Integer)
     Dim m_roundedRadius As Long, R As RECT
     
     jcTextBoxCenter = m_TextBoxHeight / 2
     
     'Draw border rectangle
     SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
     DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

     'Draw textbox border rectangle
     If m_HeaderStyle = Gradient Then
         If m_TxtBoxShadow = Shadow Then
            SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight
            OffsetRect R, 2, 2
            DrawAPIRoundRect False, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
         End If
         SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 2 - 2 * m_Indentation, m_TextBoxHeight - 1
         DrawGradientInRectangle UserControl.hdc, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor ', 3.08
     Else
         SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight
         If m_TxtBoxShadow = Shadow Then
            OffsetRect R, 2, 2
            DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
            OffsetRect R, -2, -2
         End If
         DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, m_TextBoxColor, m_FrameColor, R
     End If
    
     'set caption rect
     SetRect R_Caption, m_Indentation + m_Space * 1.5, 0, UserControl.ScaleWidth - 1 - m_Indentation - m_Space * 1.5, m_TextBoxHeight - 1
     
     'set icon coordinates
     iX = m_Indentation + m_Space * 2
     iY = (m_TextBoxHeight - m_IconSize) / 2

End Sub

Private Sub Draw_Windows(R_Caption As RECT, iY As Integer)
     Dim R As RECT, m_roundedRadius As Long
     
     jcTextBoxCenter = m_TextBoxHeight / 2
     
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    'Draw text box borders
    If m_HeaderStyle = Gradient Then
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 2, m_TextBoxHeight - 1
        DrawGradientInRectangle UserControl.hdc, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor ', 3.08
    Else
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, m_TextBoxHeight
        DrawAPIRoundRect m_RoundedCornerTxtBox, 10&, m_TextBoxColor, m_FrameColor, R
    End If
     
    'set caption rect
    SetRect R_Caption, m_Space, 0, UserControl.ScaleWidth - m_Space, m_TextBoxHeight '- 1
     
    'set icon coordinates
    iY = (m_TextBoxHeight - m_IconSize) / 2
End Sub

Private Sub Draw_Messenger(R_Caption As RECT, iY As Integer)
    Dim R As RECT, m_roundedRadius As Long
    
    jcTextBoxCenter = 0
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R
    
    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height * 2
    DrawGradientInRectangle UserControl.hdc, jcColorFrom, vbWhite, R, VerticalGradient, True, jcColorBorderPic, 2.01

    PaintShpInBar vbWhite, BlendColors(vbBlack, jcColorFrom), m_Height * 2

    If m_HeaderStyle = Gradient Or m_Enabled = False Then
        SetRect R, 0&, m_Height * 2, UserControl.ScaleWidth - 2, m_TextBoxHeight + 1
        DrawGradientInRectangle UserControl.hdc, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
        SetRect R, 0, m_Height * 2 + m_TextBoxHeight + 1, UserControl.ScaleWidth - 2, m_Height * 2 + m_TextBoxHeight + 1
        APILineEx UserControl.hdc, R.left, R.tOp, R.Right, R.Bottom, jcColorBorderPic 'vbBlack
    End If

    SetRect R, 1, 1 + m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
    DrawGradientInRectangle UserControl.hdc, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
    
    'set caption rect
    SetRect R_Caption, m_Space, m_Height * 2 + 2, UserControl.ScaleWidth - 1 - m_Space, m_TextBoxHeight + 6

    'set icon coordinates
    iY = m_Height * 2 + (m_TextBoxHeight - m_IconSize) / 2

End Sub

Private Sub Draw_InnerWedge(R_Caption As RECT, iY As Integer)
    Dim txtWidth As Integer, txtHeight As Integer, R As RECT
    Dim m_roundedRadius As Long, hFRgn As Long
    Dim poly(1 To 4) As Point, NumCoords As Long, hBrush As Long, hRgn As Long
    
    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)

    txtWidth = m_TextWidth + 10
    If txtWidth < 100 Then txtWidth = 100
    txtHeight = m_TextHeight + 5
    NumCoords = 4
    
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    If (txtWidth + R.left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.left - m_Indentation - 1
    End If
    
    'Assign values to points.
    poly(1).X = R.left:                               poly(1).y = R.tOp
    poly(2).X = R.left:                               poly(2).y = R.tOp + txtHeight
    poly(3).X = R.left + txtWidth:                    poly(3).y = R.tOp + txtHeight
    poly(4).X = R.left + txtWidth + txtHeight / 2:    poly(4).y = R.tOp
    
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.left, R.tOp, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
    
    'fill frame
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, R

    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)
    If hRgn Then FillRgn UserControl.hdc, hRgn, hBrush
    
    'draw frame borders
    APILineEx UserControl.hdc, poly(2).X, poly(2).y, poly(3).X, poly(3).y, m_FrameColor
    APILineEx UserControl.hdc, poly(3).X, poly(3).y, poly(4).X, poly(4).y, m_FrameColor
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R, True
    
    'delete created region
    DeleteObject hRgn
    DeleteObject hBrush
    
    'set caption rectangle
    SetRect R_Caption, poly(1).X + m_Indentation / 2, poly(1).y, txtWidth + poly(1).X, txtHeight + poly(1).y + 2
    
'    'set icon coordinates
'    Iy = (txtHeight - m_IconSize) / 2
    UserControl.FillStyle = 0
End Sub

Private Sub Draw_OuterWedge(R_Caption As RECT, iY As Integer)
    Dim txtWidth As Integer, txtHeight As Integer, R As RECT, r1 As RECT
    Dim m_roundedRadius As Long, hFRgn As Long
    Dim poly(1 To 4) As Point, NumCoords As Long, hBrush As Long, hRgn As Long
    
    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)

    txtWidth = m_TextWidth + 10
    If txtWidth < 100 Then txtWidth = 100
    txtHeight = m_TextHeight + 5
    NumCoords = 4
    
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    If (txtWidth + R.left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.left - m_Indentation - 1
    End If
    
    'Assign values to points.
    poly(1).X = R.left + 6:                          poly(1).y = R.tOp
    poly(2).X = R.left + 6:                          poly(2).y = R.tOp + txtHeight
    poly(3).X = R.left + txtWidth + txtHeight / 2:   poly(3).y = R.tOp + txtHeight
    poly(4).X = R.left + txtWidth:                   poly(4).y = R.tOp
    
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    
    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)
    If hRgn Then FillRgn UserControl.hdc, hRgn, hBrush
    
    'fill frame
    SetRect r1, 0&, 0&, txtWidth * 0.9, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 10&, m_TextBoxColor, m_FrameColor, r1
    SetRect r1, txtWidth * 0.9 - 5, 1, txtWidth * 0.9 + 3, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 0&, m_TextBoxColor, m_TextBoxColor, r1
    
    SetRect r1, -1, -1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, r1, True
    
   'draw frame borders
    UserControl.ForeColor = m_FrameColor
    APILineEx UserControl.hdc, poly(1).X, poly(1).y, poly(4).X, poly(4).y, UserControl.ForeColor
    APILineEx UserControl.hdc, poly(4).X, poly(4).y, poly(3).X, poly(3).y, UserControl.ForeColor
    
    RoundRect UserControl.hdc, R.left, R.tOp + txtHeight, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
    RoundRect UserControl.hdc, R.left, R.tOp + txtHeight, R.left + 10, R.tOp + txtHeight + 10, 0, 0
    
    UserControl.ForeColor = m_FillColor
    RoundRect UserControl.hdc, R.left + 1, R.tOp + txtHeight + 1, R.left + 10, R.tOp + txtHeight + 10, 0, 0
    
    'delete created region
    DeleteObject hRgn
    DeleteObject hBrush
    
    'set caption rectangle
    SetRect R_Caption, poly(1).X + m_Indentation / 2 - 6, poly(1).y, txtWidth + poly(1).X - 6, txtHeight + poly(1).y + 2
    
End Sub

Private Sub Draw_Header(R_Caption As RECT)
    Dim p_left As Long
    
    APILineEx UserControl.hdc, 0&, jcTextBoxCenter, UserControl.ScaleWidth, jcTextBoxCenter, IIf(m_Enabled, TranslateColor(&H80000015), TranslateColor(TEXT_INACTIVE)) 'TranslateColor(&H80000015)&H808080
    APILineEx UserControl.hdc, 0&, jcTextBoxCenter + 1, UserControl.ScaleWidth, jcTextBoxCenter + 1, vbWhite
    
    If Len(m_Caption) <> 0 Then

        If m_Alignment = vbLeftJustify Then
            p_left = 0 'm_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Space
        Else
            p_left = (UserControl.ScaleWidth - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hdc, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor 'TranslateColor(Ambient.BackColor)
        APILineEx UserControl.hdc, p_left, jcTextBoxCenter + 1, p_left + m_TextWidth + m_Space, jcTextBoxCenter + 1, m_FillColor 'TranslateColor(Ambient.BackColor)
        
        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If
End Sub

Private Sub Draw_Panel(R_Caption As RECT, iY As Integer)
    Dim R As RECT, m_roundedRadius As Long, hFRgn As Long, hRgn As Long
    
    jcTextBoxCenter = m_TextBoxHeight / 2
    
    'Draw border rectangle
    UserControl.FillColor = m_FillColor
    
    If m_ThemeColor = Custom Or m_HeaderStyle = TxtBoxColor Then
        UserControl.ForeColor = m_FrameColor
    Else
        UserControl.ForeColor = jcColorBorderPic
    End If

    'If m_Enabled = False Then UserControl.ForeColor = m_Border_Inactive

    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 9&)

    SetRect R, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
    If m_HeaderStyle = Gradient Then
        DrawGradientInRectangle UserControl.hdc, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, False, UserControl.ForeColor, 2.03
    End If
    
    'Creates first region to fill with color.
    hRgn = CreateRoundRectRgn(R.left, R.tOp, R.Right, R.Bottom, 0&, 0&)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.left, R.tOp, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
 
    SetWindowRgn UserControl.hwnd, hRgn, True
    
    UserControl.FillStyle = IIf(m_HeaderStyle = Gradient, 1, 0)
    
    If UserControl.ForeColor <> UserControl.BackColor Or m_HeaderStyle = TxtBoxColor Then
        RoundRect UserControl.hdc, R.left, R.tOp, R.Right - 1, R.Bottom - 1, m_roundedRadius, m_roundedRadius
        UserControl.FillStyle = 0
        DrawCorners UserControl.ForeColor
    End If
    
    'set caption rect
    SetRect R_Caption, m_Space, 0&, UserControl.ScaleWidth - m_Space, UserControl.ScaleHeight - 2

    'set icon coordinates
    iY = (UserControl.ScaleHeight - m_IconSize) / 2
End Sub

'==========================================================================
' API Functions and subroutines
'==========================================================================

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)

    'Use the API LineTo for Fast Drawing
    Dim pt As Point
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, y1, pt
    LineTo lhdcEx, x2, y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Private Function APIRectangle(ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional lColor As OLE_COLOR = -1) As Long
    'Draw an api rectangle
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As Point
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X, y, pt
    LineTo hdc, X + w, y
    LineTo hdc, X + w, y + h
    LineTo hdc, X, y + h
    LineTo hdc, X, y
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Function

Private Sub DrawGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long, Optional blnVertical = True)
    On Error Resume Next
    
    'Draw a Vertical or horizontal Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
        dR = (sR - eR) / y2
        dG = (sG - eG) / y2
        dB = (sB - eB) / y2
        For ni = 1 To y2 - 1
            APILineEx lhdcEx, X, y + ni, x2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    Else
        dR = (sR - eR) / x2
        dG = (sG - eG) / x2
        dB = (sB - eB) / x2
        For ni = 1 To x2 - 1
            APILineEx lhdcEx, X + ni, y, X + ni, y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    End If
End Sub

'Blend two colors
Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long)
    BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lColor As Long) As Long

    If OleTranslateColor(lColor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
End Function

Private Function TrimWord(strCaption As String, lngWidth As Long) As String
    Dim strCaptionAux As String, lngLenOfText As Long
    TrimWord = strCaption
    If TextWidth(strCaption) > lngWidth Then
        lngLenOfText = Len(strCaption)
        Do Until TextWidth(TrimWord & "...") <= lngWidth Or lngLenOfText = 0
            lngLenOfText = lngLenOfText - 1
            TrimWord = left(TrimWord, lngLenOfText)
        Loop
        If lngLenOfText = 0 Then TrimWord = Empty Else TrimWord = TrimWord & "..."
    End If
End Function

Private Sub EraseRegion()
    Dim hRgn As Long
    'Creates second region to fill with color.
    hRgn = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&)
    SetWindowRgn UserControl.hwnd, hRgn, True
    'delete our elliptical region
    DeleteObject hRgn
    UserControl.FillStyle = 0
End Sub

Private Sub DrawCorners(pencolor As Long)
    With UserControl
        'left top corner
         SetPixel .hdc, 0, 4, pencolor
         SetPixel .hdc, 4, 0, pencolor
         'left bottom corner
         SetPixel .hdc, .ScaleWidth - 5, 0, pencolor
         SetPixel .hdc, .ScaleWidth - 1, 4, pencolor
         'right top corner
         SetPixel .hdc, 0, .ScaleHeight - 5, pencolor
         SetPixel .hdc, 4, .ScaleHeight - 1, pencolor
         'right bottom corner
         SetPixel .hdc, .ScaleWidth - 5, .ScaleHeight - 1, pencolor
         SetPixel .hdc, .ScaleWidth - 1, .ScaleHeight - 5, pencolor
    End With
End Sub

Private Sub DrawAPIRoundRect(blnRounded As Boolean, LngRoundValue As Long, MyFillColor As Long, MyBorderColor As Long, R As RECT, Optional blnTransparent As Boolean = False)
    Dim m_roundedRadius As Long
    
    UserControl.FillColor = MyFillColor
    UserControl.ForeColor = MyBorderColor
    UserControl.FillStyle = IIf(blnTransparent = True, 1, 0)
    
    m_roundedRadius = IIf(blnRounded = False, 0&, LngRoundValue)
    
    RoundRect UserControl.hdc, R.left, R.tOp, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
    UserControl.FillStyle = 0
End Sub

Private Sub DrawGradientInRectangle(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, R As RECT, GradientType As jcGradConst, Optional blnDrawBorder As Boolean = False, Optional lBorderColor As Long = vbBlack, Optional LightCenter As Double = 2.01)
    Select Case GradientType
        Case VerticalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.left, R.tOp, R.Right + R.left, R.Bottom, True
        Case horizontalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.left, R.tOp, R.Right, R.Bottom + R.tOp, False
        Case VCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, R, True, LightCenter
        Case hCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, R, False, LightCenter
    End Select
    If blnDrawBorder Then APIRectangle lhdcEx, R.left, R.tOp, R.Right, R.Bottom, lBorderColor
End Sub

Private Sub DrawGradCilinder(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, R As RECT, Optional ByVal blnVertical As Boolean = True, Optional ByVal LightCenter As Double = 2.01)
    If LightCenter <= 1# Then LightCenter = 1.01
    If blnVertical Then
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, R.left, R.tOp, R.Right + R.left, R.Bottom / LightCenter, True
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.left, R.tOp + R.Bottom / LightCenter - 1, R.Right + R.left, (LightCenter - 1) * R.Bottom / LightCenter + 1, True
    Else
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, R.left, R.tOp, R.Right / LightCenter, R.Bottom + R.tOp, False
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.left + R.Right / LightCenter - 1, R.tOp, (LightCenter - 1) * R.Right / LightCenter + 1, R.Bottom + R.tOp, False
    End If
End Sub

Private Sub FrameEnabled(blnValor As Boolean)
    Dim c As Control
    On Error Resume Next
    For Each c In UserControl.ContainedControls: c.Enabled = blnValor: Next
End Sub

Private Sub SetDisabledColor()
    m_FrameColorDis = TranslateColor(m_Border_Inactive)
    m_TextBoxColorDis = TranslateColor(m_BtnFace) '_Inactive)
    m_FillColorDis = TranslateColor(Ambient.BackColor)
    jcColorToDis = TranslateColor(m_BtnFace_Inactive)
    jcColorFromDis = TranslateColor(m_BtnFace_Inactive)
    jcColorBorderPicDis = TranslateColor(m_Border_Inactive)
End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)

    If DstW = 0 Or DstH = 0 Then Exit Sub
    
    Dim b As Long, h As Long, F As Long, I As Long, newW As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    Dim hOldOb As Long
    Dim SrcDC As Long, tObj As Long, ttt As Long

    SrcDC = CreateCompatibleDC(hdc)

    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim hBrush As Long
        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(TransColor) 'MaskColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
        DeleteObject hBrush
    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If
    useMask = True
    If Not useMask Then TransColor = -1

    newW = DstW - 1

    For h = 0 To DstH - 1
        F = h * DstW
        For b = 0 To newW
            I = F + b
            If GetNearestColor(hdc, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
                With Data1(I)
                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                        Else
                            Data1(I) = BrushRGB
                        End If
                    Else
                        If isGreyscale Then
                            gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                        Else
                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                            Else
                                Data1(I) = Data2(I)
                            End If
                        End If
                    End If
                End With
            End If
        Next b
    Next h

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC: DeleteDC Sr2DC
    DeleteObject tObj: DeleteDC SrcDC
End Sub



'FORM TRANSPARENTE
Private Sub jcTransp(TranslucenceLevel As Byte)
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    If (osinfo.dwPlatformId = 1) Then Exit Sub
    
    SetWindowLong UserControl.Parent.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes UserControl.Parent.hwnd, 0, TranslucenceLevel, LWA_ALPHA
End Sub
