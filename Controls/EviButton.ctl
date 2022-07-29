VERSION 5.00
Begin VB.UserControl EviButton 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   LockControls    =   -1  'True
   PropertyPages   =   "EviButton.ctx":0000
   ScaleHeight     =   1470
   ScaleWidth      =   2910
   ToolboxBitmap   =   "EviButton.ctx":003E
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgHAND 
      Height          =   480
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "EviButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Evi Collection Control XP                      '
'                          By Evi Indra Effendi                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As textparametreleri) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef iccInit As ICCEX) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function IsThemeActiveAPI Lib "uxtheme.dll" Alias "IsThemeActive" () As Boolean

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Origin As Long
Private m_Stat As Long
Private m_Tats As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type ControlType
    cntrlObjectForm As Object
    cntrlHwnd As Long
    cntrlToolTipsText As String
    cntrlToolTipsTitle As String
    cntrlToolTipsIcon As Integer
End Type

Dim m_ControlType() As ControlType

Private Type TOOLINFO
    cbSize As Long
    dwFlags As Long
    hWnd As Long
    dwID As Long
    rtRect As RECT
    hInst As Long
    lpszText As Long
    lParam As Long
End Type

Private Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGB
    Red As Double
    Green As Double
    Blue As Double
End Type
Private Type ICCEX
    dwSize As Long
    dwICC As Long
End Type
'enum picture position
Public Enum EviPicturePosition
    eviTopJustify = 0
    eviLeftJustify = 1
    eviRightJustify = 2
    eviBottomJustify = 3
End Enum
'enum evi button style
Public Enum EviButtonStyle
    eviStandardButton = 0
    eviFlatButton = 1
    eviOfficeXPButton = 2
    eviWindoXPButton = 3
    eviNoBorderButton = 4
    eviWin10Button = 5
    eviWin7Button = 6
End Enum
'enum icon size
Public Enum IconSizeEnum
    [16 x 16] = 0
    [32 x 32] = 1
    [Default] = 2
    [Custom] = 3
End Enum
'client rect
Private mvarClientRect As RECT
'picture rect
Private mvarPictureRect As RECT
'caption rect
Private mvarCaptionRect As RECT
Dim mvarOrgRect As RECT
Dim g_FocusRect As RECT
Dim alan As RECT
Dim m_OriginalPicSizeW As Long
Dim m_OriginalPicSizeH As Long
Dim m_PictureOriginal As Picture
Dim m_PictureHover As Picture
Dim m_Caption As String
Dim m_PicturePosition As EviPicturePosition
Dim m_ButtonStyle As EviButtonStyle
Dim m_Picture As Picture
Dim m_PictureWidth As Long
Dim m_PictureHeight As Long
Dim m_PictureSize As IconSizeEnum
Dim mvarDrawTextParams As textparametreleri
Dim g_HasFocus As Byte
Dim g_MouseDown As Byte, g_MouseIn As Byte
Attribute g_MouseIn.VB_VarUserMemId = 1073938455
Dim g_Button As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Attribute g_Button.VB_VarUserMemId = 1073938457
Attribute g_Shift.VB_VarUserMemId = 1073938457
Attribute g_X.VB_VarUserMemId = 1073938457
Attribute g_Y.VB_VarUserMemId = 1073938457
Dim g_KeyPressed As Byte
Attribute g_KeyPressed.VB_VarUserMemId = 1073938461
Dim m_ShowFocusRect As Boolean
Attribute m_ShowFocusRect.VB_VarUserMemId = 1073938462
Dim WithEvents g_Font As StdFont
Attribute g_Font.VB_VarHelpID = -1
Const mvarPadding As Byte = 4
Dim m_BEVEL As Integer
Attribute m_BEVEL.VB_VarUserMemId = 1073938463
Dim m_BEVELDEPTH As Integer
Attribute m_BEVELDEPTH.VB_VarUserMemId = 1073938464
Dim m_TransparentBG As Boolean
Attribute m_TransparentBG.VB_VarUserMemId = 1073938465
Dim m_MaskColor As OLE_COLOR
Attribute m_MaskColor.VB_VarUserMemId = 1073938466
Dim m_XPShowBorderAlways As Boolean
Attribute m_XPShowBorderAlways.VB_VarUserMemId = 1073938467
Dim m_DefCurHand As Boolean
Attribute m_DefCurHand.VB_VarUserMemId = 1073938468
Dim m_ForeColor As OLE_COLOR
Attribute m_ForeColor.VB_VarUserMemId = 1073938469
Dim m_BackColor As OLE_COLOR
Attribute m_BackColor.VB_VarUserMemId = 1073938470
Dim m_OSStyle As Boolean
Dim m_XPDefaultColors As Boolean
Attribute m_XPDefaultColors.VB_VarUserMemId = 1073938471
Dim m_XPColor_Pressed As OLE_COLOR
Attribute m_XPColor_Pressed.VB_VarUserMemId = 1073938472
Dim m_XPColor_Hover As OLE_COLOR
Attribute m_XPColor_Hover.VB_VarUserMemId = 1073938473
'for tool tip
Private m_Object As Object
Attribute m_Object.VB_VarUserMemId = 1073938474
Dim m_ToolTipText As String
Attribute m_ToolTipText.VB_VarUserMemId = 1073938475
Dim m_ToolTipTitle As String
Attribute m_ToolTipTitle.VB_VarUserMemId = 1073938476
Dim m_ToolTipIcon As ttIconType
Attribute m_ToolTipIcon.VB_VarUserMemId = 1073938477
Dim m_Counter As Long
Attribute m_Counter.VB_VarUserMemId = 1073938478
Private ghWndTip As Long, ghWndParent As Long
Attribute ghWndTip.VB_VarUserMemId = 1073938479
Attribute ghWndParent.VB_VarUserMemId = 1073938479
'
Private anmCnt As Long
Private Const totalAnmCnt As Long = 15

Enum ttIconType
    [No Icon] = 0
    [Icon Info] = 1
    [Icon Warning] = 2
    [Icon Error] = 3
End Enum

Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Const ICC_WIN95_CLASSES As Long = &HFF

Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETWINDOWTHEME As Long = (CCM_FIRST + &HB)
Private Const WM_USER As Long = &H400
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const ECM_FIRST As Long = &H1500

Private Const EM_SHOWBALLOONTIP = ECM_FIRST + 3

Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOPMOST As Long = &H8&

Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"

Private Const TTF_ABSOLUTE As Long = &H80
Private Const TTF_CENTERTIP As Long = &H2
Private Const TTF_DI_SETITEM As Long = &H8000
Private Const TTF_IDISHWND As Long = &H1
Private Const TTF_RTLREADING As Long = &H4
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_TRACK As Long = &H20
Private Const TTF_TRANSPARENT As Long = &H100

Private Const TTI_ERROR As Long = 3
Private Const TTI_INFO As Long = 1
Private Const TTI_NONE As Long = 0
Private Const TTI_WARNING As Long = 2

Private Const TTM_ACTIVATE As Long = (WM_USER + 1)
Private Const TTM_ADDTOOL As Long = (WM_USER + 4)
Private Const TTM_ADJUSTRECT As Long = (WM_USER + 31)
Private Const TTM_DELTOOL As Long = (WM_USER + 5)
Private Const TTM_ENUMTOOLS As Long = (WM_USER + 14)
Private Const TTM_GETBUBBLESIZE As Long = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOL As Long = (WM_USER + 15)
Private Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Private Const TTM_GETMARGIN As Long = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH As Long = (WM_USER + 25)
Private Const TTM_GETTEXT As Long = (WM_USER + 11)
Private Const TTM_GETTIPBKCOLOR As Long = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR As Long = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Private Const TTM_GETTOOLINFO As Long = (WM_USER + 8)
Private Const TTM_HITTEST As Long = (WM_USER + 10)
Private Const TTM_NEWTOOLRECT As Long = (WM_USER + 6)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_POPUP As Long = (WM_USER + 34)
Private Const TTM_RELAYEVENT As Long = (WM_USER + 7)
Private Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Private Const TTM_SETMARGIN As Long = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETTITLE As Long = (WM_USER + 32)
Private Const TTM_SETTOOLINFO As Long = (WM_USER + 9)
Private Const TTM_SETWINDOWTHEME As Long = CCM_SETWINDOWTHEME
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTM_TRACKPOSITION As Long = (WM_USER + 18)
Private Const TTM_UPDATE As Long = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Private Const TTM_WINDOWFROMPOINT As Long = (WM_USER + 16)

Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFO As Long = (TTN_FIRST - 0)
Private Const TTN_LAST As Long = (-549)
Private Const TTN_LINKCLICK As Long = (TTN_FIRST - 3)
Private Const TTN_NEEDTEXT As Long = TTN_GETDISPINFO
Private Const TTN_POP As Long = (TTN_FIRST - 2)
Private Const TTN_SHOW As Long = (TTN_FIRST - 1)

Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_NOANIMATE As Long = &H10
Private Const TTS_NOFADE As Long = &H20
Private Const TTS_NOPREFIX As Long = &H2

'declare event
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)

Private Sub UserControl_InitProperties()
    On Error GoTo error

    m_BackColor = &H8000000F
    m_ForeColor = &H80000012
    m_ShowFocusRect = 1
    Set UserControl.Font = Ambient.Font
    Set g_Font = Ambient.Font
    m_Caption = Ambient.DisplayName
    m_PicturePosition = 1
    m_PictureWidth = 32
    m_PictureHeight = 32
    m_PictureSize = 1
    Set m_PictureHover = LoadPicture("")
    Set m_PictureOriginal = LoadPicture("")
    m_XPColor_Pressed = &H80000014
    m_XPColor_Hover = &H80000016
    m_XPDefaultColors = 1
    m_OSStyle = True
    m_ButtonStyle = GetButtonStyleFromOs

    m_DefCurHand = 0
    m_XPShowBorderAlways = 0
    m_MaskColor = 0
    m_TransparentBG = 0
    m_BEVEL = 1
    m_BEVELDEPTH = 8

    Set m_Object = UserControl.Parent

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "InitProperties : " & Err.Description
End Sub

Private Sub UserControl_Paint()
    On Error GoTo error

    Set m_Object = UserControl.Parent

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "Paint : " & Err.Description
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo error
    Set m_Object = UserControl.Parent
    m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = m_BackColor
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.ForeColor = m_ForeColor
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", UserControl.Extender.ToolTipText)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", 0)
    UserControl.Extender.ToolTipText = PropBag.ReadProperty("ToolTipText", m_ToolTipText)
    m_ShowFocusRect = PropBag.ReadProperty("Focus", 1)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_PicturePosition = PropBag.ReadProperty("IconPosition", 1)
    Set m_Picture = PropBag.ReadProperty("Icon", Nothing)
    m_PictureWidth = PropBag.ReadProperty("IconWidth", 32)
    m_PictureHeight = PropBag.ReadProperty("IconHeight", 32)
    m_PictureSize = PropBag.ReadProperty("IconSize", 1)
    m_OriginalPicSizeW = PropBag.ReadProperty("OriginalPicSizeW", 32)
    m_OriginalPicSizeH = PropBag.ReadProperty("OriginalPicSizeH", 32)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_PictureHover = PropBag.ReadProperty("IconHover", Nothing)
    Set m_PictureOriginal = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)

    m_OSStyle = PropBag.ReadProperty("ButtonStyleOS", True)

    If m_OSStyle Then
        m_ButtonStyle = GetButtonStyleFromOs
    Else
        m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", 2)
    End If

    m_XPColor_Pressed = PropBag.ReadProperty("ColorPressed", &H80000014)
    m_XPColor_Hover = PropBag.ReadProperty("ColorHover", &H80000016)
    m_XPDefaultColors = PropBag.ReadProperty("DefaultColors", 1)

    m_DefCurHand = PropBag.ReadProperty("DefCurHand", 0)
    m_XPShowBorderAlways = PropBag.ReadProperty("ShowBorder", 0)
    m_MaskColor = PropBag.ReadProperty("MaskColor", 0)
    m_TransparentBG = PropBag.ReadProperty("Transparent", 0)
    m_BEVEL = PropBag.ReadProperty("BEVEL", 1)
    m_BEVELDEPTH = PropBag.ReadProperty("BEVELDEPTH", 8)
    SetAccessKeys

    UserControl_Resize
error:
End Sub

Private Sub UserControl_Show()
    On Error GoTo error

    Set m_Object = UserControl.Parent
    ShowTool
error:
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo error

    Set g_Font = Nothing

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "Terminate : " & Err.Description
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("IconPosition", m_PicturePosition, 1)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, 2)
    Call PropBag.WriteProperty("ButtonStyleOS", m_OSStyle, True)
    Call PropBag.WriteProperty("Icon", m_Picture, Nothing)
    Call PropBag.WriteProperty("IconWidth", m_PictureWidth, 32)
    Call PropBag.WriteProperty("IconHeight", m_PictureHeight, 32)
    Call PropBag.WriteProperty("IconSize", m_PictureSize, 1)
    Call PropBag.WriteProperty("OriginalPicSizeW", m_OriginalPicSizeW, 32)
    Call PropBag.WriteProperty("OriginalPicSizeH", m_OriginalPicSizeH, 32)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, UserControl.Extender.ToolTipText)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, 0)
    Call PropBag.WriteProperty("IconHover", m_PictureHover, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Focus", m_ShowFocusRect, 1)
    Call PropBag.WriteProperty("ColorPressed", m_XPColor_Pressed, &H80000014)
    Call PropBag.WriteProperty("ColorHover", m_XPColor_Hover, &H80000016)
    Call PropBag.WriteProperty("DefaultColors", m_XPDefaultColors, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("DefCurHand", m_DefCurHand, 0)
    Call PropBag.WriteProperty("ShowBorder", m_XPShowBorderAlways, 0)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, 0)
    Call PropBag.WriteProperty("Transparent", m_TransparentBG, 0)
    Call PropBag.WriteProperty("BEVEL", m_BEVEL, 1)
    Call PropBag.WriteProperty("BEVELDEPTH", m_BEVELDEPTH, 8)
End Sub
Private Sub CalcRECTs()
    On Error GoTo error
    Dim picWidth, picHeight, capWidth, capHeight As Long
    With alan
        .Left = 0
        .Top = 0
        .Right = ScaleWidth - 1
        .Bottom = ScaleHeight - 1
    End With

    With mvarClientRect
        .Left = alan.Left + mvarPadding
        .Top = alan.Top + mvarPadding
        .Right = alan.Right - mvarPadding + 1
        .Bottom = alan.Bottom - mvarPadding + 1
    End With

    If m_Picture Is Nothing Then
        With mvarCaptionRect
            .Left = mvarClientRect.Left
            .Top = mvarClientRect.Top
            .Right = mvarClientRect.Right
            .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
    Else
        If m_Caption = "" Then
            With mvarPictureRect
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - m_PictureWidth) \ 2) + mvarClientRect.Left
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - m_PictureHeight) \ 2) + mvarClientRect.Top
                .Right = mvarPictureRect.Left + m_PictureWidth
                .Bottom = mvarPictureRect.Top + m_PictureHeight
            End With
            Exit Sub
        End If

        With mvarCaptionRect
            .Left = mvarClientRect.Left
            .Top = mvarClientRect.Top
            .Right = mvarClientRect.Right
            .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect

        picWidth = m_PictureWidth
        picHeight = m_PictureHeight
        capWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
        capHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top


        If m_PicturePosition = 1 Then
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
                .Left = mvarPictureRect.Right + mvarPadding
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With

        ElseIf m_PicturePosition = 2 Then
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
                .Left = mvarCaptionRect.Right + mvarPadding
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
        ElseIf m_PicturePosition = 0 Then
            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
            With mvarCaptionRect
                .Top = mvarPictureRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
        ElseIf m_PicturePosition = 3 Then
            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarCaptionRect.Top + capHeight
                .Right = mvarCaptionRect.Left + capWidth
            End With
            With mvarPictureRect
                .Top = mvarCaptionRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
                .Bottom = mvarPictureRect.Top + picHeight
                .Right = mvarPictureRect.Left + picWidth
            End With
        End If
    End If

error:
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo error
    Set g_Font = New StdFont

    ScaleMode = 3
    PaletteMode = 3

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "Initialize : " & Err.Description
End Sub

' Returns the version of Windows that the user is running
Public Function GetButtonStyleFromOs() As EviButtonStyle
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetButtonStyleFromOs = eviStandardButton

            Case VER_PLATFORM_WIN32_NT
                GetButtonStyleFromOs = eviStandardButton

                Select Case osv.dwVerMajor
                    Case 3
                        GetButtonStyleFromOs = eviStandardButton
                    Case 4
                        GetButtonStyleFromOs = eviStandardButton
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetButtonStyleFromOs = eviStandardButton
                            Case 1
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWindoXPButton
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If

                            Case 2
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWindoXPButton
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If

                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWin7Button
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If
                            Case 1
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWin7Button
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If
                            Case 2
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWin10Button
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If
                            Case 3
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWin10Button
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If
                            Case Else
                                If IsThemeActiveAPI Then
                                    GetButtonStyleFromOs = eviWin10Button
                                Else
                                    GetButtonStyleFromOs = eviStandardButton
                                End If
                        End Select
                    Case Else
                        If IsThemeActiveAPI Then
                            GetButtonStyleFromOs = eviWin10Button
                        Else
                            GetButtonStyleFromOs = eviStandardButton
                        End If
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetButtonStyleFromOs = eviStandardButton
                    Case 90
                        GetButtonStyleFromOs = eviStandardButton
                    Case Else
                        GetButtonStyleFromOs = eviStandardButton
                End Select
        End Select
    Else
        GetButtonStyleFromOs = eviStandardButton
    End If
End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If Not Me.Enabled Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Refresh
End Sub

Private Sub UserControl_EnterFocus()
    g_HasFocus = 1
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    g_HasFocus = 0
    g_MouseDown = 0
    'anmCnt = totalAnmCnt
    Refresh
End Sub

Private Sub UserControl_Resize()
    On Error GoTo error

    If ScaleWidth < 10 Then UserControl.Width = 150
    If ScaleHeight < 10 Then UserControl.Height = 150

    m_Stat = ScaleWidth
    m_Tats = ScaleHeight

    g_FocusRect.Left = 2
    g_FocusRect.Right = ScaleWidth - 2
    g_FocusRect.Top = 2
    g_FocusRect.Bottom = ScaleHeight - 2

    DeleteObject Origin

    If m_ButtonStyle = eviWindoXPButton Or m_ButtonStyle = eviWin7Button Then
        RoundCorners
    End If

    Refresh

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "Resize : " & Err.Description
End Sub
Public Sub Refresh()
    On Error GoTo error

    AutoRedraw = True

    UserControl.Cls

    XPAdjustColorScheme

    If m_ButtonStyle <> eviNoBorderButton Then Draw3DEffect

    CalcRECTs
    DrawPicture

    If g_HasFocus = 1 And m_ShowFocusRect And m_ButtonStyle <> eviWindoXPButton And m_ButtonStyle <> eviWin7Button And m_ButtonStyle <> eviWin10Button Then DrawFocusRect hdc, g_FocusRect

    DrawCaption

    AutoRedraw = False
error:
End Sub

Private Sub ShowTool()
    On Error GoTo error

    'For owner form focus loose
    Exit Sub

    Set m_Object = UserControl.Parent
    m_ToolTipText = UserControl.Extender.ToolTipText
    UserControl.Extender.ToolTipText = ""
    AddToolTipText m_Object, hWnd, m_ToolTipText, m_ToolTipTitle, m_ToolTipIcon
    ShowToolTipText
error:
End Sub

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_Text As String)
    m_ToolTipText = New_Text
    PropertyChanged "ToolTipText"
End Property

'Public Property Get ToolTipTitle() As String
'    ToolTipTitle = m_ToolTipTitle
'End Property
'
'Public Property Let ToolTipTitle(ByVal New_Title As String)
'    m_ToolTipTitle = New_Title
'    PropertyChanged "ToolTipTitle"
'End Property

'Public Property Get ToolTipIcon() As ttIconType
'    ToolTipIcon = m_ToolTipIcon
'End Property
'
'Public Property Let ToolTipIcon(ByVal New_Icon As ttIconType)
'    m_ToolTipIcon = New_Icon
'    PropertyChanged "ToolTipIcon"
'End Property

Private Sub UserControl_DblClick()
    SetCapture hWnd
    UserControl_MouseDown g_Button, g_Shift, g_X, g_Y
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If g_KeyPressed = 0 Then


        If KeyCode = 32 Then
            g_MouseDown = 1
            'anmCnt = totalAnmCnt
            g_MouseIn = 1
            Refresh
        End If
        g_KeyPressed = 1
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        g_MouseDown = 0
        g_MouseIn = 0
        'anmCnt = totalAnmCnt
        Refresh

        UserControl_MouseUp 1, Shift, 0, 0
    End If
    g_KeyPressed = 0
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    g_Button = Button: g_Shift = Shift: g_X = X: g_Y = Y
    If Button <> 2 Then
        'anmCnt = totalAnmCnt
        g_MouseDown = 1
        Refresh
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
        If g_MouseIn = 0 Then
            OverTimer.Enabled = True
            'anmCnt = totalAnmCnt
            g_MouseIn = 1
            If Not m_PictureHover Is Nothing Then
                Set m_Picture = m_PictureHover
            End If
            RaiseEvent MouseIn(Shift)
            Refresh
            DoEvents

        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    g_MouseDown = 0
    'anmCnt = totalAnmCnt

    If Button <> 2 Then
        Refresh
        If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
            RaiseEvent Click
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get Font() As Font
    Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    With g_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    PropertyChanged "Font"
End Property

Private Sub g_Font_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = g_Font
    Refresh
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get Focus() As Boolean
    Focus = m_ShowFocusRect
End Property

Public Property Let Focus(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    PropertyChanged "Focus"
    Refresh
End Property

Private Sub RunXTRA3D(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
    On Error GoTo error
    Dim T As Integer
    Dim TEMPRENK As Long

    TEMPRENK = RENK
    BEVELDEPTHH = BEVELDEPTHH * (-1)

    For T = BEVELL To 0 Step -1
        TEMPRENK = COLOR_DarkenLightenColor(TEMPRENK, BEVELDEPTHH)
        DRAWRECT hdc, T, T, ScaleWidth - T, ScaleHeight - T, TEMPRENK, 0
    Next T

    BEVELDEPTHH = BEVELDEPTHH * (-1)

    For T = BEVELL To 0 Step -1
        RENK = RGB(COLOR_LongToRGB(RENK).Red + BEVELDEPTHH, COLOR_LongToRGB(RENK).Green + BEVELDEPTHH, COLOR_LongToRGB(RENK).Blue + BEVELDEPTHH)
        DrawLine T, T, ScaleWidth - (T + 1), T, RENK
        DrawLine T, T, T, ScaleHeight - (T + 1), RENK
    Next T
error:
End Sub
Private Sub RunXTRA3D_PRESSED(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
    Dim ret As Integer
    Dim GRIN As Integer
    Dim BLU As Integer
    Dim T As Integer
    On Error GoTo error
    Dim TEMPRENK As Long
    TEMPRENK = RENK

    For T = BEVELL To 0 Step -1
        ret = COLOR_LongToRGB(TEMPRENK).Red + BEVELDEPTHH
        GRIN = COLOR_LongToRGB(TEMPRENK).Green + BEVELDEPTHH
        BLU = COLOR_LongToRGB(TEMPRENK).Blue + BEVELDEPTHH
        TEMPRENK = RGB(ret, GRIN, BLU)
        DRAWRECT hdc, T, T, ScaleWidth - T, ScaleHeight - T, TEMPRENK, 0
    Next T


    BEVELDEPTHH = BEVELDEPTHH * (-1)
    For T = BEVELL To 0 Step -1
        RENK = COLOR_DarkenLightenColor(RENK, BEVELDEPTHH)
        DrawLine T, T, ScaleWidth - (T + 1), T, RENK
        DrawLine T, T, T, ScaleHeight - (T + 1), RENK
    Next T
error:
End Sub
Private Sub RunShowBorderOnFocus(RENK As Long, BEVELL As Integer, BEVELDEPTHH As Integer)
    Dim T As Integer
    On Error GoTo error

    If BEVELL < 2 Then
        DRAWRECT hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, &H80000010
        DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
        DRAWRECT hdc, -1, -1, ScaleWidth + 1, ScaleHeight + 1, &H80000015
    Else
        RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, BEVELDEPTH + 3
    End If
error:
End Sub
Private Sub XPAdjustColorScheme()
    On Error GoTo error

    Dim Hedef As RGB
    Dim KOLOR As RGB

    If m_ButtonStyle = eviWindoXPButton Or m_ButtonStyle = eviWin7Button Then Exit Sub

    If m_ButtonStyle = eviOfficeXPButton Then
        If m_TransparentBG = True And g_MouseDown = 0 Then
            Transparentia
        Else
            If DefaultColors = True Then
                UserControl.BackColor = vbButtonFace
            Else
                UserControl.BackColor = m_BackColor
            End If
        End If
    ElseIf m_ButtonStyle = eviWin10Button And g_MouseDown = 0 And g_MouseIn = 0 Then
        If DefaultColors = True Then
            UserControl.BackColor = RGB(225, 225, 225)
        Else
            UserControl.BackColor = m_BackColor
        End If
    Else
        If m_TransparentBG = True And g_MouseDown = 0 Then
            Transparentia
        End If
    End If

    If m_ButtonStyle = eviOfficeXPButton Then
        If DefaultColors = True Then
            UserControl.BackColor = RGB(237, 239, 241)
        Else
            UserControl.BackColor = m_BackColor
        End If

        If g_MouseDown = 0 And g_MouseIn = 1 Then
            If DefaultColors = True Then
                UserControl.BackColor = RGB(254, 241, 189)
            Else
                UserControl.BackColor = ColorHover
            End If
        End If

        If g_MouseDown = 1 Then
            If DefaultColors = True Then
                UserControl.BackColor = RGB(255, 229, 117)
            Else
                UserControl.BackColor = ColorPressed
            End If
        End If
    End If

    If m_ButtonStyle = eviWin10Button Then
        If g_MouseDown = 0 And g_MouseIn = 1 Then
            If DefaultColors = True Then
                UserControl.BackColor = RGB(229, 241, 251)
                Exit Sub
            Else
                UserControl.BackColor = ColorHover
                Exit Sub
            End If
        End If

        If g_MouseDown = 1 Then
            If DefaultColors = True Then
                UserControl.BackColor = RGB(204, 228, 247)
                Exit Sub
            Else
                UserControl.BackColor = ColorPressed
                Exit Sub
            End If
        End If

        If DefaultColors = True Then
            UserControl.BackColor = RGB(225, 225, 225)
        Else
            UserControl.BackColor = m_BackColor
        End If
    End If
error:
End Sub

Private Sub Draw3DEffect()
    On Error GoTo error

    If Not Ambient.UserMode Then
        If m_ButtonStyle = eviWindoXPButton Then
            DrawWinXPButton 0
        ElseIf m_ButtonStyle = eviWin7Button Then
            DrawWin7Button 0
        ElseIf m_ButtonStyle = eviOfficeXPButton Then
            XPAdjustColorScheme
        ElseIf m_ButtonStyle = eviWin10Button Then
            XPAdjustColorScheme
        Else
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, m_BEVELDEPTH
            End If
        End If
        Exit Sub
    End If

    If m_ButtonStyle = eviOfficeXPButton Then
        If Not (ShowBorder = False And g_MouseIn = 0) Then
            DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, m_ForeColor
        End If
    ElseIf m_ButtonStyle = eviWin10Button Then
        If g_MouseDown = 0 Then
            If g_MouseIn = 0 Then
                If g_HasFocus <> 0 Then
                    DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, RGB(0, 120, 215)
                    DRAWRECT hdc, 1, 1, ScaleWidth - 1, ScaleHeight - 1, RGB(0, 120, 215)
                Else
                    DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, RGB(173, 173, 173)
                End If
            Else
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, RGB(0, 120, 215)
            End If
        Else
            DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, RGB(0, 84, 153)
        End If
    ElseIf m_ButtonStyle = eviWindoXPButton Then
        If g_MouseDown = 1 Then DrawWinXPButton 2
        If g_MouseDown = 0 And g_MouseIn = 1 Then DrawWinXPButton 0, 1
        If g_MouseDown = 0 And g_MouseIn = 0 Then DrawWinXPButton 0
    ElseIf m_ButtonStyle = eviWin7Button Then
        If g_MouseDown = 1 Then DrawWin7Button 2
        If g_MouseDown = 0 And g_MouseIn = 1 Then DrawWin7Button 0, 1
        If g_MouseDown = 0 And g_MouseIn = 0 Then DrawWin7Button 0
    Else
        If g_MouseDown = 1 Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000014
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000010
            Else
                RunXTRA3D_PRESSED COLOR_UniColor(UserControl.BackColor), m_BEVEL, m_BEVELDEPTH
            End If
        End If

        If g_MouseDown = 0 And g_MouseIn = 1 Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, m_BEVELDEPTH
            End If
        End If

        If g_MouseDown = 0 And g_MouseIn = 0 And m_ButtonStyle = eviStandardButton Then
            If m_BEVEL < 2 Then
                DRAWRECT hdc, 0, 0, ScaleWidth, ScaleHeight, &H80000010
                DRAWRECT hdc, 0, 0, ScaleWidth + 1, ScaleHeight + 1, &H80000014
            Else
                RunXTRA3D COLOR_UniColor(UserControl.BackColor), m_BEVEL, m_BEVELDEPTH
            End If
        End If

        If (g_HasFocus = 1 And m_ButtonStyle = eviStandardButton And g_MouseDown = 0) Then
            RunShowBorderOnFocus COLOR_UniColor(UserControl.BackColor), m_BEVEL, m_BEVELDEPTH
        End If

        If (g_HasFocus = 1 And m_ButtonStyle = eviWin10Button And g_MouseDown = 0) Then
            RunShowBorderOnFocus RGB(0, 120, 215), 0, 2
        End If
    End If

    Exit Sub
error:
    Debug.Print Err.Description
End Sub

Private Sub OverTimer_Timer()
    On Error GoTo error

    Dim P As POINTAPI

    GetCursorPos P

    If hWnd <> WindowFromPoint(P.X, P.Y) Then
        OverTimer.Enabled = False

        g_MouseIn = 0
        RaiseEvent MouseOut(g_Shift)
        Refresh

        'simulate mouse away
        If g_MouseDown = 1 Then
            g_MouseDown = 0
            Refresh
            g_MouseDown = 1
        End If
    End If

    Exit Sub
error:
    DebugPrint "[" & UserControl.Name & "] " & "Timer : " & Err.Description
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    SetAccessKeys
    Refresh
End Property
Public Property Get ButtonStyle() As EviButtonStyle
    ButtonStyle = m_ButtonStyle
End Property
Public Property Let ButtonStyle(ByVal New_ButtonStyle As EviButtonStyle)
    If Not m_OSStyle Then
        m_ButtonStyle = New_ButtonStyle
        PropertyChanged "ButtonStyle"

        If m_ButtonStyle = eviWindoXPButton Or m_ButtonStyle = eviWin7Button Then Transparent = False

        If Not m_XPDefaultColors And m_ButtonStyle = eviWin7Button Then
            m_XPColor_Hover = COLOR_DarkenLightenColor(m_BackColor, 15, False)
            m_XPColor_Pressed = COLOR_DarkenLightenColor(m_BackColor, -25, False)
        End If

        UserControl_Resize
    End If
End Property

Public Property Get IconPosition() As EviPicturePosition
    IconPosition = m_PicturePosition
End Property
Public Property Let IconPosition(ByVal New_PicturePosition As EviPicturePosition)
    m_PicturePosition = New_PicturePosition
    PropertyChanged "IconPosition"
    Refresh
End Property
Public Property Get Icon() As Picture
    Set Icon = m_Picture
End Property
Public Property Set Icon(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set m_PictureOriginal = New_Picture
    If m_Picture Is Nothing Then
        m_OriginalPicSizeW = 32
        m_OriginalPicSizeH = 32
    Else
        m_OriginalPicSizeW = UserControl.ScaleX(m_Picture.Width, 8, UserControl.ScaleMode)
        m_OriginalPicSizeH = UserControl.ScaleY(m_Picture.Height, 8, UserControl.ScaleMode)
    End If
    PropertyChanged "Icon"
    If m_PictureSize = 2 Then
        m_PictureWidth = UserControl.ScaleX(m_Picture.Width, 8, UserControl.ScaleMode)
        m_PictureHeight = UserControl.ScaleY(m_Picture.Height, 8, UserControl.ScaleMode)
    End If
    Refresh
End Property

Public Property Get IconWidth() As Long
    IconWidth = m_PictureWidth
End Property
Public Property Let IconWidth(ByVal New_PictureWidth As Long)
    m_PictureWidth = New_PictureWidth
    PropertyChanged "IconWidth"
    Refresh
End Property
Public Property Get IconHeight() As Long
    IconHeight = m_PictureHeight
End Property
Public Property Let IconHeight(ByVal New_PictureHeight As Long)
    m_PictureHeight = New_PictureHeight
    PropertyChanged "IconHeight"
    Refresh
End Property
Public Property Get IconSize() As IconSizeEnum
    IconSize = m_PictureSize
End Property
Public Property Let IconSize(ByVal New_PictureSize As IconSizeEnum)
    m_PictureSize = New_PictureSize
    PropertyChanged "IconSize"

    If New_PictureSize = 0 Then
        m_PictureWidth = 16
        m_PictureHeight = 16
    ElseIf New_PictureSize = 1 Then
        m_PictureWidth = 32
        m_PictureHeight = 32
    ElseIf New_PictureSize = 2 Then
        If Not (m_Picture Is Nothing) Then
            m_PictureWidth = m_OriginalPicSizeW
            m_PictureHeight = m_OriginalPicSizeH
        Else
            m_PictureWidth = 32
            m_PictureHeight = 32
        End If
    End If

    Refresh
End Property

Private Sub CalculateCaptionRect()
    On Error GoTo error
    Dim mvarWidth, mvarHeight As Long
    Dim mvarFormat As Long
    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)
    End With
    mvarFormat = &H400 Or &H10 Or &H4 Or &H1
    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, mvarFormat, mvarDrawTextParams
    mvarWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
    mvarHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
    With mvarCaptionRect
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (mvarCaptionRect.Right - mvarCaptionRect.Left)) \ 2)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (mvarCaptionRect.Bottom - mvarCaptionRect.Top)) \ 2)
        .Right = mvarCaptionRect.Left + mvarWidth
        .Bottom = mvarCaptionRect.Top + mvarHeight
    End With
error:
End Sub

Private Sub DrawCaption()
    On Error GoTo error
    If m_Caption = "" Then Exit Sub

    SetTextColor hdc, COLOR_UniColor(m_ForeColor)

    Dim mvarForeColor As OLE_COLOR
    mvarOrgRect = mvarCaptionRect

    If g_MouseDown = 1 And m_ButtonStyle <> 2 And m_ButtonStyle <> eviWin10Button And m_ButtonStyle <> eviWin7Button Then
        With mvarCaptionRect
            .Left = mvarCaptionRect.Left + 1
            .Top = mvarCaptionRect.Top + 1
            .Right = mvarCaptionRect.Right + 1
            .Bottom = mvarCaptionRect.Bottom + 1
        End With
    End If

    If Not Enabled Then
        Dim g_tmpFontColor As OLE_COLOR
        g_tmpFontColor = UserControl.ForeColor

        If m_ButtonStyle <> eviWin7Button And m_ButtonStyle <> eviWin10Button Then
            SetTextColor hdc, COLOR_UniColor(&H80000014)
            Dim mvarCaptionRect_Iki As RECT
            With mvarCaptionRect_Iki
                .Bottom = mvarCaptionRect.Bottom
                .Left = mvarCaptionRect.Left + 1
                .Right = mvarCaptionRect.Right + 1
                .Top = mvarCaptionRect.Top + 1
            End With
            DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, &H10 Or &H4 Or &H1, mvarDrawTextParams
        End If

        SetTextColor hdc, COLOR_UniColor(&H80000010)
        DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams

        SetTextColor hdc, COLOR_UniColor(g_tmpFontColor)
        Exit Sub
    End If

    DrawTextEx hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
    mvarCaptionRect = mvarOrgRect
error:
End Sub
Private Sub DrawBitmap(EnabledPic As Byte, CurPictRECT As RECT, Optional AsShadow As Byte = 0)
    Dim DC1 As Long
    Dim BM1 As Long
    Dim DC2 As Long
    Dim BM2 As Long
    Dim UZUN1 As Long
    Dim UZUN2 As Long
    Dim hBrush As Long
    On Error GoTo error
    DC1 = CreateCompatibleDC(hdc)
    DC2 = CreateCompatibleDC(hdc)
    BM1 = CreateCompatibleBitmap(hdc, m_OriginalPicSizeW, m_OriginalPicSizeH)
    BM2 = CreateCompatibleBitmap(hdc, m_PictureWidth, m_PictureHeight)
    UZUN1 = SelectObject(DC1, BM1)
    UZUN2 = SelectObject(DC2, BM2)

    If EnabledPic = 0 Then
        Dim DC3 As Long
        Dim BM3 As Long

        DC3 = CreateCompatibleDC(hdc)
        BM3 = SelectObject(DC3, m_Picture.Handle)

        SetBkColor DC1, &HFFFFFF

        DRAWRECT DC1, 0, 0, _
                 m_OriginalPicSizeW, m_OriginalPicSizeH, &HFFFFFF, 1

        TransParentPic DC1, DC1, DC3, 0, 0, _
                       m_OriginalPicSizeW, m_OriginalPicSizeH, 0, 0, m_MaskColor

        StretchBlt DC2, 0, 0, _
                   m_PictureWidth, _
                   m_PictureHeight, _
                   DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020

        SelectObject DC2, UZUN2

        If AsShadow = 1 Then
            hBrush = CreateSolidBrush(RGB(146, 146, 146))
            Call DrawState(hdc, hBrush, 0, BM2, 0, CurPictRECT.Left, _
                           CurPictRECT.Top, 0, 0, &H80& Or &H4&)
            DeleteObject hBrush
        Else
            Call DrawState(hdc, 0, 0, BM2, 0, CurPictRECT.Left, _
                           CurPictRECT.Top, 0, 0, &H20& Or &H4&)
        End If

        DeleteObject BM3
        DeleteDC DC3

    Else
        Call DrawState(DC1, 0, 0, m_Picture, 0, 0, 0, 0, 0, _
                       &H0 Or &H4&)

        StretchBlt DC2, 0, 0, _
                   m_PictureWidth, _
                   m_PictureHeight, _
                   DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020

        TransParentPic hdc, hdc, DC2, 0, 0, _
                       m_PictureWidth, m_PictureHeight, _
                       CurPictRECT.Left, CurPictRECT.Top, m_MaskColor

    End If
    SelectObject DC1, UZUN1
    SelectObject DC2, UZUN2
    DeleteObject BM1
    DeleteObject BM2
    DeleteDC DC1
    DeleteDC DC2
error:
End Sub



Private Sub DrawPIcon(EnabledPic As Byte, CurPictRECT As RECT, Optional AsShadow As Byte = 0)
    On Error GoTo error
    If EnabledPic = 0 Then
        Dim DC1 As Long
        Dim BM1 As Long
        Dim DC2 As Long
        Dim BM2 As Long
        Dim UZUN1 As Long
        Dim UZUN2 As Long
        Dim hBrush As Long

        DC1 = CreateCompatibleDC(hdc)
        BM1 = CreateCompatibleBitmap(hdc, m_OriginalPicSizeW, m_OriginalPicSizeH)

        DC2 = CreateCompatibleDC(hdc)
        BM2 = CreateCompatibleBitmap(hdc, m_PictureWidth, m_PictureHeight)

        UZUN1 = SelectObject(DC1, BM1)
        UZUN2 = SelectObject(DC2, BM2)

        If AsShadow = 1 Then
            hBrush = CreateSolidBrush(RGB(146, 146, 146))
            Call DrawState(DC1, hBrush, 0, m_Picture, 0, 0, 0, 0, 0, _
                           &H80& Or &H3&)
            DeleteObject hBrush
        Else
            Call DrawState(DC1, 0, 0, m_Picture, 0, 0, 0, 0, 0, _
                           &H20& Or &H3&)
        End If

        StretchBlt DC2, 0, 0, _
                   CurPictRECT.Right - CurPictRECT.Left, _
                   CurPictRECT.Bottom - CurPictRECT.Top, _
                   DC1, 0, 0, m_OriginalPicSizeW, m_OriginalPicSizeH, &HCC0020

        TransParentPic hdc, hdc, DC2, 0, 0, _
                       m_PictureWidth, m_PictureHeight, _
                       CurPictRECT.Left, CurPictRECT.Top, &H0

        SelectObject DC1, UZUN1
        SelectObject DC2, UZUN2
        DeleteObject BM1
        DeleteObject BM2
        DeleteDC DC1
        DeleteDC DC2

    Else
        UserControl.PaintPicture m_Picture, CurPictRECT.Left, _
                                 CurPictRECT.Top, CurPictRECT.Right - CurPictRECT.Left, _
                                 CurPictRECT.Bottom - CurPictRECT.Top, 0, 0, _
                                 m_OriginalPicSizeW, m_OriginalPicSizeH
    End If
error:
End Sub

Private Sub DrawPicture()
    On Error GoTo error

    Dim Margin As Integer

    If m_Picture Is Nothing Then Exit Sub

    mvarOrgRect = mvarPictureRect

    If g_MouseDown = 0 And g_MouseIn = 1 And (m_ButtonStyle = eviOfficeXPButton) Then
        Margin = -3
    ElseIf g_MouseDown = 1 And Not (m_ButtonStyle = eviOfficeXPButton) Then
        Margin = 1
    End If

    With mvarPictureRect
        .Left = .Left + Margin
        .Top = .Top + Margin
        .Right = .Right + Margin
        .Bottom = .Bottom + Margin
    End With

    If m_Picture.Type = 1 Then
        If Not Enabled Then
            DrawBitmap 0, mvarPictureRect
        Else
            If g_MouseDown = 0 And g_MouseIn = 1 And (m_ButtonStyle = eviOfficeXPButton) Then
                DrawBitmap 0, mvarOrgRect, 1
            End If

            DrawBitmap 1, mvarPictureRect
        End If
    ElseIf m_Picture.Type = 3 Then
        If Not Enabled Then
            DrawPIcon 0, mvarPictureRect
        Else
            If g_MouseDown = 0 And g_MouseIn = 1 And (m_ButtonStyle = eviOfficeXPButton) Then _
               DrawPIcon 0, mvarOrgRect, 1

            DrawPIcon 1, mvarPictureRect
        End If
    End If

    mvarPictureRect = mvarOrgRect

error:
End Sub
Private Sub Transparentia()
    On Error Resume Next
    Dim RESIM As StdPicture
    Dim mem_dc As Long
    Dim mem_bm As Long
    Dim orig_bm As Long
    Dim wid As Long
    Dim hgt As Long
    Dim IX As Long
    Dim YE As Long

    IX = ScaleX(Extender.Left, Parent.ScaleMode, ScaleMode)
    YE = ScaleY(Extender.Top, Parent.ScaleMode, ScaleMode)

    Set RESIM = Parent.Picture
    mem_dc = CreateCompatibleDC(hdc)
    mem_bm = CreateCompatibleBitmap(mem_dc, ScaleWidth, ScaleHeight)

    SelectObject mem_dc, RESIM.Handle

    BitBlt hdc, 0, 0, ScaleWidth, ScaleHeight, _
           mem_dc, IX, YE, &HCC0020

    SelectObject mem_dc, orig_bm
    DeleteObject mem_bm
    DeleteDC mem_dc
    Set RESIM = Nothing
End Sub

Public Property Get IconHover() As Picture
    Set IconHover = m_PictureHover
End Property

Public Property Set IconHover(ByVal New_PictureHover As Picture)
    Set m_PictureHover = New_PictureHover
    PropertyChanged "IconHover"
End Property
Public Property Get ColorPressed() As OLE_COLOR
    ColorPressed = m_XPColor_Pressed
End Property

Public Property Let ColorPressed(ByVal New_XPColor_Pressed As OLE_COLOR)
    m_XPColor_Pressed = New_XPColor_Pressed
    PropertyChanged "ColorPressed"
End Property
Public Property Get ColorHover() As OLE_COLOR
    ColorHover = m_XPColor_Hover
End Property

Public Property Let ColorHover(ByVal New_XPColor_Hover As OLE_COLOR)
    m_XPColor_Hover = New_XPColor_Hover
    PropertyChanged "ColorHover"
End Property
Public Property Get DefaultColors() As Boolean
    DefaultColors = m_XPDefaultColors
End Property

Public Property Let DefaultColors(ByVal New_XPDefaultColors As Boolean)
    m_XPDefaultColors = New_XPDefaultColors

    If Not New_XPDefaultColors Then
        m_XPColor_Hover = COLOR_DarkenLightenColor(m_BackColor, 10, False)
        m_XPColor_Pressed = COLOR_DarkenLightenColor(m_BackColor, -25, False)
    End If

    Refresh

    PropertyChanged "DefaultColors"
End Property

Public Property Get ButtonStyleOS() As Boolean
    ButtonStyleOS = m_OSStyle
End Property

Public Property Let ButtonStyleOS(ByVal New_OSStyle As Boolean)
    If New_OSStyle Then
        ButtonStyle = GetButtonStyleFromOs()
    End If

    m_OSStyle = New_OSStyle

    PropertyChanged "ButtonStyleOS"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor

    If Not m_XPDefaultColors Then
        m_XPColor_Hover = COLOR_DarkenLightenColor(m_BackColor, 10, False)
        m_XPColor_Pressed = COLOR_DarkenLightenColor(m_BackColor, -25, False)
    End If

    Refresh
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
    Refresh
End Property
Public Property Get DefCurHand() As Boolean
    DefCurHand = m_DefCurHand
End Property

Public Property Let DefCurHand(ByVal New_DefCurHand As Boolean)
    m_DefCurHand = New_DefCurHand
    PropertyChanged "DefCurHand"
End Property

Public Property Get ShowBorder() As Boolean
    ShowBorder = m_XPShowBorderAlways
End Property

Public Property Let ShowBorder(ByVal New_XPShowBorderAlways As Boolean)
    m_XPShowBorderAlways = New_XPShowBorderAlways
    PropertyChanged "ShowBorder"
End Property
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    Refresh
End Property
Public Property Get Transparent() As Boolean
    Transparent = m_TransparentBG
End Property

Public Property Let Transparent(ByVal New_TransparentBG As Boolean)
    m_TransparentBG = New_TransparentBG
    PropertyChanged "Transparent"
    Refresh
End Property

'Public Property Get BEVEL() As Integer
'    BEVEL = m_BEVEL
'End Property
'
'Public Property Let BEVEL(ByVal New_BEVEL As Integer)
'    m_BEVEL = New_BEVEL
'    PropertyChanged "BEVEL"
'    Refresh
'End Property
'Public Property Get BEVELDEPTH() As Integer
'    BEVELDEPTH = m_BEVELDEPTH
'End Property
'
'Public Property Let BEVELDEPTH(ByVal New_BEVELDEPTH As Integer)
'    m_BEVELDEPTH = New_BEVELDEPTH
'    PropertyChanged "BEVELDEPTH"
'    Refresh
'End Property

Private Function COLOR_LongToRGB(UniColorValue As Long) As RGB
    Dim BlueS As Double, GreenS As Double, RGBs As String
    COLOR_LongToRGB.Blue = Fix((UniColorValue / 256) / 256)
    BlueS = (COLOR_LongToRGB.Blue * 256) * 256
    COLOR_LongToRGB.Green = Fix((UniColorValue - BlueS) / 256)
    GreenS = COLOR_LongToRGB.Green * 256
    COLOR_LongToRGB.Red = Fix(UniColorValue - BlueS - GreenS)

End Function
Private Function COLOR_UniColor(ColorVal As Long) As Long

    COLOR_UniColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then COLOR_UniColor = GetSysColor(ColorVal And &HFFFFFF)
End Function
Private Function COLOR_DarkenLightenColor(ByVal Color As Long, ByVal Value As Long, Optional Luminosity As Boolean = True) As Long
    Dim R As Long, G As Long, b As Long

    b = ((Color \ &H10000) Mod &H100)

    If Luminosity Then
        b = b + ((b * Value) \ &HC0)
    Else
        b = b + Value
    End If

    G = ((Color \ &H100) Mod &H100) + Value
    R = (Color And &HFF) + Value

    If R < 0 Then R = 0
    If R > 255 Then R = 255
    If G < 0 Then G = 0
    If G > 255 Then G = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255

    COLOR_DarkenLightenColor = RGB(R, G, b)
End Function

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim pt As POINTAPI

    Dim pen As Long: pen = CreatePen(0, 1, Color)
    Dim ret As Long: ret = SelectObject(hdc, pen)

    If ret <> 0 Then
        MoveToEx hdc, X1, Y1, pt
        LineTo hdc, X2, Y2
    End If

    DeleteObject pen
End Sub

Private Sub DRAWRECT(DestHDC As Long, ByVal RectLEFT As Long, _
                     ByVal RectTOP As Long, _
                     ByVal RectRIGHT As Long, ByVal RectBOTTOM As Long, _
                     ByVal MyColor As Long, _
                     Optional FillRectWithColor As Byte = 0)
    Dim MyRect As RECT, Firca As Long
    Firca = CreateSolidBrush(COLOR_UniColor(MyColor))
    With MyRect
        .Left = RectLEFT
        .Top = RectTOP
        .Right = RectRIGHT
        .Bottom = RectBOTTOM
    End With
    If FillRectWithColor = 1 Then FillRect DestHDC, MyRect, Firca Else FrameRect DestHDC, MyRect, Firca
    DeleteObject Firca
End Sub

Private Sub DrawWinXPButton(ByVal None_Press_Disabled As Byte, Optional HOVERING As Byte)
    Dim X As Long, Intg As Single, curBackColor As Long, OuterBorderColor As Long
    Dim KolorHover As Long, KolorPressed As Long, KolorBack As Long

    OuterBorderColor = &H80000015

    If Enabled Then
        If m_XPDefaultColors = True Then
            KolorBack = &H8000000F
        Else
            KolorBack = m_BackColor
            'KolorPressed = m_XPColor_Pressed
            'KolorHover = m_XPColor_Hover
        End If

        KolorPressed = RGB(140, 170, 230)
        KolorHover = RGB(225, 153, 71)

        DRAWRECT hdc, 0, 0, m_Stat, m_Tats, KolorBack, 1

        If None_Press_Disabled = 0 Then
            Intg = 25 / m_Tats: curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(KolorBack), 48)
            For X = 1 To m_Tats
                DrawLine 0, X, m_Stat, X, COLOR_DarkenLightenColor(curBackColor, -Intg * X)
            Next

            DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
            SetPixel hdc, 1, 1, OuterBorderColor
            SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
            SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
            SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

            If HOVERING <> 1 Then
                If g_HasFocus = 1 Then
                    DRAWRECT hdc, 1, 2, m_Stat - 1, m_Tats - 2, KolorPressed
                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), -33)
                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 65)
                    DrawLine 1, 2, m_Stat - 1, 2, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 50)
                    DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 31)
                    DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(KolorPressed), 31)
                Else
                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -48)
                    DrawLine 1, m_Tats - 3, m_Stat - 2, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, -32)
                    DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -36)
                    DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, -24)
                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(curBackColor, 16)
                    DrawLine 1, 2, m_Stat - 2, 2, COLOR_DarkenLightenColor(curBackColor, 10)
                    DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -5)
                    DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, -10)
                End If
            Else
                DRAWRECT hdc, 1, 2, m_Stat - 1, m_Tats - 2, KolorHover
                DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(KolorHover, -40)
                DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(KolorHover, 90)
                DrawLine 1, 2, m_Stat - 1, 2, COLOR_DarkenLightenColor(KolorHover, 35)
                DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(KolorHover, 20)
                DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(KolorHover, 20)
            End If
        ElseIf None_Press_Disabled = 2 Then
            Intg = 15 / m_Tats
            curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(KolorBack), 48)
            curBackColor = COLOR_DarkenLightenColor(curBackColor, -32)
            For X = 1 To m_Tats
                DrawLine 0, m_Tats - X, m_Stat, m_Tats - X, COLOR_DarkenLightenColor(curBackColor, -Intg * X)
            Next
            DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
            SetPixel hdc, 1, 1, OuterBorderColor
            SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
            SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
            SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

            DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, 16)
            DrawLine 1, m_Tats - 3, m_Stat - 2, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, 10)
            DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, 5)
            DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, curBackColor
            DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(curBackColor, -32)
            DrawLine 1, 2, m_Stat - 2, 2, COLOR_DarkenLightenColor(curBackColor, -24)
            DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -32)
            DrawLine 2, 2, 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -22)
        End If
    Else
        curBackColor = COLOR_DarkenLightenColor(COLOR_UniColor(KolorBack), 48)
        DRAWRECT hdc, 0, 0, m_Stat, m_Tats, COLOR_DarkenLightenColor(curBackColor, -24), 1
        DRAWRECT hdc, 0, 0, m_Stat, m_Tats, COLOR_DarkenLightenColor(curBackColor, -84)
        SetPixel hdc, 1, 1, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, 1, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, m_Stat - 2, 1, COLOR_DarkenLightenColor(curBackColor, -72)
        SetPixel hdc, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, -72)
    End If
End Sub

Public Sub Gradient(ColorStart As Long, ColorEnd As Long, _
                    StartX As Long, StartY As Long, EndX As Long, EndY As Long, _
                    Horizonal As Boolean)

    Dim R, G, b As Single
    Dim Wi%

    'Obj.AutoRedraw = True
    'Obj.ScaleMode = 3
    Dim StartColor As RGB
    Dim EndColor As RGB

    StartColor = COLOR_LongToRGB(ColorStart)
    EndColor = COLOR_LongToRGB(ColorEnd)

    If Not Horizonal Then
        R = (EndColor.Red - StartColor.Red) / (EndY - StartY)
        G = (EndColor.Green - StartColor.Green) / (EndY - StartY)
        b = (EndColor.Blue - StartColor.Blue) / (EndY - StartY)

        For Wi = 0 To EndY - StartY - 1
            DrawLine StartX, StartY + Wi, EndX, StartY + Wi, COLOR_UniColor(RGB(StartColor.Red, StartColor.Green, StartColor.Blue))

            StartColor.Red = StartColor.Red + R
            StartColor.Green = StartColor.Green + G
            StartColor.Blue = StartColor.Blue + b
        Next Wi
    Else
        R = (EndColor.Red - StartColor.Red) / (EndX - StartX)
        G = (EndColor.Green - StartColor.Green) / (EndX - StartX)
        b = (EndColor.Blue - StartColor.Blue) / (EndX - StartX)

        For Wi = 0 To EndX - StartX - 1
            DrawLine StartX + Wi, StartY, StartX + Wi, EndY, COLOR_UniColor(RGB(StartColor.Red, StartColor.Green, StartColor.Blue))

            StartColor.Red = StartColor.Red + R
            StartColor.Green = StartColor.Green + G
            StartColor.Blue = StartColor.Blue + b
        Next Wi
    End If
End Sub

Private Sub DrawWin7Button(ByVal None_Press_Disabled As Byte, Optional HOVERING As Byte)
    Dim dng As Long, Intg As Single, curBackColor As Long, OuterBorderColor As Long
    Dim KolorHover As Long, KolorPressed As Long
    Dim middle As Long
    Dim adim As Double
    Dim StartColor As Long
    Dim EndColor As Long

    If Enabled Then
        If m_XPDefaultColors = True Then
            KolorPressed = RGB(140, 170, 230)
            KolorHover = RGB(158, 176, 186)
            curBackColor = RGB(242, 242, 242)
        Else
            KolorPressed = m_XPColor_Pressed
            KolorHover = m_XPColor_Hover
            curBackColor = m_BackColor
        End If

        ' Not pressed
        If None_Press_Disabled = 0 Then
            ' Normal
            If HOVERING <> 1 Then
                If m_XPDefaultColors = True Then
                    middle = CLng(m_Tats * 0.5)

                    StartColor = COLOR_UniColor(RGB(242, 242, 242))
                    EndColor = COLOR_UniColor(RGB(235, 235, 235))

                    Call Gradient(StartColor, EndColor, 0, 1, m_Stat, middle, False)

                    StartColor = COLOR_UniColor(RGB(221, 221, 221))
                    EndColor = COLOR_UniColor(RGB(207, 207, 207))

                    Call Gradient(StartColor, EndColor, 0, middle, m_Stat, m_Tats, False)
                Else
                    middle = CLng(m_Tats * 0.5)
                    adim = 15 / middle

                    For dng = 0 To middle
                        DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(curBackColor, -adim * dng)
                    Next

                    curBackColor = COLOR_DarkenLightenColor(curBackColor, -25)

                    For dng = middle To m_Tats
                        DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(curBackColor, -adim * (dng - middle))
                    Next
                End If

                If g_HasFocus = 1 Then
                    OuterBorderColor = RGB(60, 127, 177)
                    DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
                    SetPixel hdc, 1, 1, OuterBorderColor
                    SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
                    SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
                    SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

                    'DRAWRECT hdc, 1, 2, m_Stat - 1, m_Tats - 2, RGB(64, 215, 252)

                    'Alt
                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -40)
                    DrawLine 2, m_Tats - 3, m_Stat - 2, m_Tats - 3, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -10)
                    'Sað
                    DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -40)
                    DrawLine m_Stat - 3, 2, m_Stat - 3, m_Tats - 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -10)
                    'Üst
                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -40)
                    DrawLine 2, 2, m_Stat - 2, 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -10)
                    'Sol
                    DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -40)
                    DrawLine 2, 2, 2, m_Tats - 2, COLOR_DarkenLightenColor2(RGB(64, 215, 252), -10)
                    '                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                    '                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                    '                    DrawLine 1, 2, m_Stat - 1, 2, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                    '                    DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                    '                    DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                Else
                    OuterBorderColor = RGB(112, 112, 112)
                    DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
                    SetPixel hdc, 1, 1, OuterBorderColor
                    SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
                    SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
                    SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

                    'Alt
                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, 10)
                    'Sað
                    DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, 10)
                    'Üst
                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(curBackColor, 10)
                    'Sol
                    DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor(curBackColor, 10)

                    '                    DrawLine 1, m_Tats - 3, m_Stat - 2, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, 10)
                    '                    DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, 10)
                    '                    DrawLine 1, 2, m_Stat - 2, 2, COLOR_DarkenLightenColor(curBackColor, 10)
                    '                    DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(curBackColor, 10)
                End If
            Else        ' Hover
                If m_XPDefaultColors = True Then
                    middle = CLng(m_Tats * 0.5)

                    StartColor = COLOR_UniColor(RGB(234, 246, 253))
                    EndColor = COLOR_UniColor(RGB(217, 240, 252))

                    Call Gradient(StartColor, EndColor, 0, 1, m_Stat, middle, False)

                    StartColor = COLOR_UniColor(RGB(190, 230, 253))
                    EndColor = COLOR_UniColor(RGB(167, 217, 245))

                    Call Gradient(StartColor, EndColor, 0, middle, m_Stat, m_Tats, False)
                Else
                    middle = CLng(m_Tats * 0.5)
                    adim = 15 / middle

                    For dng = 0 To middle
                        DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(KolorHover, -adim * dng)
                    Next

                    curBackColor = COLOR_DarkenLightenColor(KolorHover, -25)

                    For dng = middle To m_Tats
                        DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(curBackColor, -adim * (dng - middle))
                    Next
                End If

                'Border
                OuterBorderColor = RGB(60, 127, 177)
                DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
                SetPixel hdc, 1, 1, OuterBorderColor
                SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
                SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
                SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

                'Alt
                DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(RGB(234, 246, 253), 10)
                'Sað
                DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(RGB(234, 246, 253), 10)
                'Üst
                DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(RGB(234, 246, 253), 10)
                'Sol
                DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor(RGB(234, 246, 253), 10)

                '                If g_HasFocus = 1 Then
                '                    DRAWRECT hdc, 1, 2, m_Stat - 1, m_Tats - 2, RGB(64, 215, 252)
                '                    DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                '                    DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                '                    DrawLine 1, 2, m_Stat - 1, 2, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                '                    DrawLine 2, 3, 2, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                '                    DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(COLOR_UniColor(RGB(64, 215, 252)), 11)
                '                 End If

            End If
        ElseIf None_Press_Disabled = 2 Then        ' Pressed
            If m_XPDefaultColors = True Then
                middle = CLng(m_Tats * 0.5)

                StartColor = COLOR_UniColor(RGB(229, 244, 252))
                EndColor = COLOR_UniColor(RGB(196, 229, 246))

                Call Gradient(StartColor, EndColor, 0, 1, m_Stat, middle, False)

                middle = m_Tats - middle

                StartColor = COLOR_UniColor(RGB(152, 209, 239))
                EndColor = COLOR_UniColor(RGB(104, 179, 219))

                Call Gradient(StartColor, EndColor, 0, CLng(m_Tats * 0.5), m_Stat, m_Tats, False)
            Else
                middle = CLng(m_Tats * 0.5)
                adim = 15 / middle

                curBackColor = COLOR_DarkenLightenColor(KolorHover, -10)

                For dng = 0 To middle
                    DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(KolorHover, -adim * dng)
                Next

                curBackColor = COLOR_DarkenLightenColor(KolorPressed, -15)

                For dng = middle To m_Tats
                    DrawLine 0, dng, m_Stat, dng, COLOR_DarkenLightenColor(curBackColor, -adim * (dng - middle))
                Next
            End If

            OuterBorderColor = RGB(44, 98, 139)
            DRAWRECT hdc, 0, 0, m_Stat, m_Tats, OuterBorderColor
            SetPixel hdc, 1, 1, OuterBorderColor
            SetPixel hdc, 1, m_Tats - 2, OuterBorderColor
            SetPixel hdc, m_Stat - 2, 1, OuterBorderColor
            SetPixel hdc, m_Stat - 2, m_Tats - 2, OuterBorderColor

            DrawLine 2, m_Tats - 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            DrawLine m_Stat - 2, 2, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            DrawLine 2, 1, m_Stat - 2, 1, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            DrawLine 1, 2, 1, m_Tats - 2, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            '            DrawLine 1, m_Tats - 3, m_Stat - 2, m_Tats - 3, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            '            DrawLine m_Stat - 3, 3, m_Stat - 3, m_Tats - 3, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            '            DrawLine 1, 2, m_Stat - 2, 2, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
            '            DrawLine 2, 2, 2, m_Tats - 2, COLOR_DarkenLightenColor(RGB(158, 176, 186), -22)
        End If
    Else        ' Button disabled
        curBackColor = COLOR_DarkenLightenColor2(COLOR_UniColor(m_BackColor), 48)
        DRAWRECT hdc, 0, 0, m_Stat, m_Tats, COLOR_DarkenLightenColor2(curBackColor, -24), 1
        DRAWRECT hdc, 0, 0, m_Stat, m_Tats, COLOR_DarkenLightenColor2(curBackColor, -84)
        SetPixel hdc, 1, 1, COLOR_DarkenLightenColor2(curBackColor, -72)
        SetPixel hdc, 1, m_Tats - 2, COLOR_DarkenLightenColor2(curBackColor, -72)
        SetPixel hdc, m_Stat - 2, 1, COLOR_DarkenLightenColor2(curBackColor, -72)
        SetPixel hdc, m_Stat - 2, m_Tats - 2, COLOR_DarkenLightenColor2(curBackColor, -72)
    End If
End Sub

Private Sub RoundCorners()
    Dim Alan1 As Long, Alan2 As Long
    Origin = CreateRectRgn(0, 0, m_Stat, m_Tats)
    Alan2 = CreateRectRgn(0, 0, 0, 0)
    Alan1 = CreateRectRgn(0, 0, 2, 1)
    CombineRgn Alan2, Origin, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(0, m_Tats, 2, m_Tats - 1)
    CombineRgn Origin, Alan2, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(m_Stat, 0, m_Stat - 2, 1)
    CombineRgn Alan2, Origin, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(m_Stat, m_Tats, m_Stat - 2, m_Tats - 1)
    CombineRgn Origin, Alan2, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(0, 1, 1, 2)
    CombineRgn Alan2, Origin, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(0, m_Tats - 1, 1, m_Tats - 2)
    CombineRgn Origin, Alan2, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(m_Stat, 1, m_Stat - 1, 2)
    CombineRgn Alan2, Origin, Alan1, 4
    DeleteObject Alan1
    Alan1 = CreateRectRgn(m_Stat, m_Tats - 1, m_Stat - 1, m_Tats - 2)
    CombineRgn Origin, Alan2, Alan1, 4
    DeleteObject Alan1
    DeleteObject Alan2
    SetWindowRgn hWnd, Origin, True
    DeleteObject Origin
End Sub
Private Sub TransParentPic(DestDC As Long, _
                           DestDCTrans As Long, _
                           SrcDC As Long, _
                           SrcRectLeft As Long, SrcRectTop As Long, _
                           SrcRectRight As Long, SrcRectBottom As Long, _
                           DstX As Long, _
                           DstY As Long, _
                           MaskColor As Long)

    Dim nRet As Long, W As Integer, H As Integer
    Dim MonoMaskDC As Long, hMonoMask As Long
    Dim MonoInvDC As Long, hMonoInv As Long
    Dim ResultDstDC As Long, hResultDst As Long
    Dim ResultSrcDC As Long, hResultSrc As Long
    Dim hPrevMask As Long, hPrevInv As Long
    Dim hPrevSrc As Long, hPrevDst As Long
    Dim SrcRect As RECT

    With SrcRect
        .Left = SrcRectLeft
        .Top = SrcRectTop
        .Right = SrcRectRight
        .Bottom = SrcRectBottom
    End With

    W = SrcRectRight - SrcRectLeft
    H = SrcRectBottom - SrcRectTop

    MonoMaskDC = CreateCompatibleDC(DestDCTrans)
    MonoInvDC = CreateCompatibleDC(DestDCTrans)
    hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)

    ResultDstDC = CreateCompatibleDC(DestDCTrans)
    ResultSrcDC = CreateCompatibleDC(DestDCTrans)
    hResultDst = CreateCompatibleBitmap(DestDCTrans, W, H)
    hResultSrc = CreateCompatibleBitmap(DestDCTrans, W, H)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)

    Dim OldBC As Long
    OldBC = SetBkColor(SrcDC, MaskColor)
    nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, _
                  SrcRect.Left, SrcRect.Top, &HCC0020)
    MaskColor = SetBkColor(SrcDC, OldBC)

    nRet = BitBlt(MonoInvDC, 0, 0, W, H, _
                  MonoMaskDC, 0, 0, &H330008)

    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                  DestDCTrans, DstX, DstY, &HCC0020)

    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                  MonoMaskDC, 0, 0, &H8800C6)


    nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, _
                  SrcRect.Left, SrcRect.Top, &HCC0020)

    nRet = BitBlt(ResultSrcDC, 0, 0, W, H, _
                  MonoInvDC, 0, 0, &H8800C6)

    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                  ResultSrcDC, 0, 0, &H660046)

    nRet = BitBlt(DestDC, DstX, DstY, W, H, _
                  ResultDstDC, 0, 0, &HCC0020)

    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask

    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv

    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst

    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc

    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC
End Sub

Private Sub SetAccessKeys()
    Dim ampersandPos As Long
    If Len(m_Caption) > 1 Then
        ampersandPos = InStr(1, m_Caption, "&", vbTextCompare)
        If (ampersandPos < Len(m_Caption)) And (ampersandPos > 0) Then
            If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
            Else
                ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)
                If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then
                    UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
                Else
                    UserControl.AccessKeys = ""
                End If
            End If
        Else
            UserControl.AccessKeys = ""
        End If
    Else
        UserControl.AccessKeys = ""
    End If
End Sub

Private Sub ShowToolTipText()
    Dim m_m As Long
    On Error GoTo error
    If m_Counter <= 0 Then Exit Sub

    For m_m = 1 To m_Counter
        ShowToolTipsBalloon m_ControlType(m_m).cntrlObjectForm, m_ControlType(m_m).cntrlHwnd, m_ControlType(m_m).cntrlToolTipsText, m_ControlType(m_m).cntrlToolTipsTitle, m_ControlType(m_m).cntrlToolTipsIcon
    Next m_m
error:
End Sub

Private Function AddToolTipText(Optional ObjectFormOwner As Object = Nothing, Optional AddObjectToShowToolTips As Long, Optional ToolTipText As String _
                                                                                                                        = "", Optional ToolTipTitle As String = "", Optional _
                                                                                                                                                                    ToolTipIcon As ttIconType = 1)
    On Error GoTo error
    m_Counter = m_Counter + 1
    ReDim Preserve m_ControlType(m_Counter)
    Set m_ControlType(m_Counter).cntrlObjectForm = ObjectFormOwner
    m_ControlType(m_Counter).cntrlHwnd = AddObjectToShowToolTips
    m_ControlType(m_Counter).cntrlToolTipsText = ToolTipText
    m_ControlType(m_Counter).cntrlToolTipsTitle = ToolTipTitle
    m_ControlType(m_Counter).cntrlToolTipsIcon = ToolTipIcon
error:
End Function

Private Sub ShowToolTipsBalloon(Optional ObjectForm As Object, Optional OwnHwnd As Long, Optional _
                                                                                         ToolTipsText As String, Optional ToolTipTitle As String, Optional _
                                                                                                                                                  ToolTipIcon As Integer)
    Dim tiInfo As TOOLINFO
    Dim MyHwnD As Long
    Dim hWndTip As Long, dwFlags As Long, ICEx As ICCEX
    On Error GoTo error
    dwFlags = TTS_NOPREFIX Or TTS_ALWAYSTIP Or TTS_BALLOON

    With ICEx
        .dwSize = Len(ICEx)
        .dwICC = ICC_WIN95_CLASSES
    End With

    InitCommonControlsEx ICEx

    hWndTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, "", WS_POPUP Or dwFlags, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, OwnHwnd, 0, App.hInstance, ByVal 0&)

    If hWndTip = 0 Then Exit Sub

    SetWindowPos hWndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    ghWndTip = hWndTip
    ghWndParent = ObjectForm.hWnd

    With tiInfo
        .dwFlags = TTF_SUBCLASS Or TTF_TRANSPARENT
        .hWnd = OwnHwnd
        .lpszText = StrPtr(StrConv(ToolTipsText, vbFromUnicode))
        .hInst = App.hInstance
        GetClientRect OwnHwnd, .rtRect

        .cbSize = Len(tiInfo)

    End With

    SendMessage ghWndTip, TTM_ADDTOOL, 0&, tiInfo
    If ToolTipTitle <> vbNullString Or ToolTipIcon <> 0 Then
        SendMessage ghWndTip, TTM_SETTITLE, CLng(ToolTipIcon), ByVal ToolTipTitle
    End If
error:
End Sub

Private Function GetRGBFromOLE(ByVal Color As OLE_COLOR) As RGB
    Dim lngColour As Long
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    Dim result As RGB

    lngColour = Color

    If lngColour < 0 Then        'Check if it is a system colour
        'get the index into the colour table
        lngColour = lngColour And Not &H80000000
        lngColour = GetSysColor(lngColour)
    End If

    result.Red = lngColour And 255
    result.Green = lngColour \ 256 And 255
    result.Blue = lngColour \ 65536 And 255

    GetRGBFromOLE = result
End Function

Private Function Red(ColourValue As Long) As Long
    Dim strHexColour As String

    strHexColour = Right$("000000" & Hex(ColourValue), 6)
    Red = CLng("&H" & Right$(strHexColour, 2))
End Function

Private Function Green(ColourValue As Long) As Long
    Dim strHexColour As String

    strHexColour = Right$("000000" & Hex(ColourValue), 6)
    Green = CLng("&H" & Mid$(strHexColour, 3, 2))
End Function

Private Function Blue(ColourValue As Long) As Long
    Dim strHexColour As String

    strHexColour = Right$("000000" & Hex(ColourValue), 6)
    Blue = CLng("&H" & Left$(strHexColour, 2))
End Function

Private Function COLOR_DarkenLightenColor2(ByVal Color As Long, ByVal Value As Double) As Long
    Dim R As Double, G As Double, b As Double

    Dim bRgb As RGB

    bRgb = COLOR_LongToRGB(COLOR_UniColor(Color))

    b = bRgb.Blue
    G = bRgb.Green
    R = bRgb.Red

    If b > G And b > R Then
        G = G + (G / b * Value)
        R = R + (R / b * Value)
        b = b + Value
    ElseIf G > b And G > R Then
        b = b + (b / G * Value)
        R = R + (R / G * Value)
        G = G + Value
    Else
        b = b + (b / R * Value)
        G = G + (G / R * Value)
        R = R + Value
    End If

    If R < 0 Then R = 0
    If R > 255 Then R = 255
    If G < 0 Then G = 0
    If G > 255 Then G = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255

    COLOR_DarkenLightenColor2 = RGB(R, G, b)
End Function
