VERSION 5.00
Begin VB.UserControl EviTextBox 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   PropertyPages   =   "EviTextBox.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   2025
   ToolboxBitmap   =   "EviTextBox.ctx":003E
   Begin VB.TextBox TXT 
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   3
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TXT 
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   2
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TXT 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   1
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   195
      Left            =   1890
      ScaleHeight     =   135
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   1965
      TabIndex        =   1
      Top             =   195
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox TXT 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   20
      TabIndex        =   0
      Top             =   20
      Width           =   1935
   End
End
Attribute VB_Name = "EviTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Evi Text Box Style XP"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Const ColorXPRec = 16777215

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private TR As RECT
Private TBR As RECT
Private m_hDC As Long
Private m_ThDC As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_hBmp As Long
Private m_hBmpOld As Long
Private xMax As Long
Private m_MemDC As Boolean
Private xLocked As Boolean
Private xBackColor As OLE_COLOR
Private xForeColor As OLE_COLOR
Private m_Color As OLE_COLOR
Private xPassword As Variant
Private xTExt As String
Private xTMP As String
Private xAlign As AlignmentConstants
Private xFont As Font
Private xScroll As ScrollBarConstants
Private xTempText As String
Private xEnabled As Boolean
Private xIcon As Picture
Private xPointer As MousePointerConstants
Private xFocus As Boolean

Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub DoTextBoxStyler()
    On Error GoTo error
    GetClientRect UserControl.hWnd, TR
    DrawFillRectangle TR, ColorXPRec, m_hDC
    DrawingObjectTextBox
    If m_MemDC Then
        With UserControl
            pDraw .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
        End With
    End If
error:
End Sub

Private Sub DrawingObjectTextBox()
    On Error GoTo error
    DrawRectangle TR, ShiftColorXP(m_Color, 100), m_hDC
    With TBR
        .Left = 1
        .Top = 1
        .Bottom = TR.Bottom - 1
        .Right = TR.Left + (TR.Right - TR.Left) * 0
    End With
    DrawFillRectangle TBR, ShiftColorXP(m_Color, 180), m_hDC
error:
End Sub

Private Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Private Sub DrawRectangle(ByRef bRect As RECT, ByVal Color As Long, ByVal hdc As Long)
    Dim hBrush As Long
    On Error GoTo error
    hBrush = CreateSolidBrush(Color)
    FrameRect hdc, bRect, hBrush
    DeleteObject hBrush
error:
End Sub

Private Sub DrawLine( _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal cHdc As Long, _
        ByVal Color As Long)
    Dim Pen1 As Long
    Dim Pen2 As Long
    Dim POS As POINTAPI
    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    MoveToEx cHdc, X, Y, POS
    LineTo cHdc, Width, Height
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
End Sub

Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
    Dim R As Long, G As Long, b As Long, Delta As Long
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    b = ((MyColor \ &H10000) Mod &H100)
    Delta = &HFF - Base
    b = Base + b * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If b > 255 Then b = 255
    ShiftColorXP = R + 256& * G + 65536 * b
End Function

Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, hRect, hBrush
    DeleteObject hBrush
End Sub

Private Function ThDC(Width As Long, Height As Long) As Long
    If m_ThDC = 0 Then
        If (Width + Height) > 0 Then pCreate Width, Height
    Else
        If Width > m_lWidth Or Height > m_lHeight Then pCreate Width, Height
    End If
    ThDC = m_ThDC
End Function

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
    Dim lhDCC As Long
    pDestroy
    lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If lhDCC Then
        m_ThDC = CreateCompatibleDC(lhDCC)
        If m_ThDC Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
            If m_hBmp Then
                m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
                If m_hBmpOld Then
                    m_lWidth = Width
                    m_lHeight = Height
                    DeleteDC lhDCC
                    Exit Sub
                End If
            End If
        End If
        DeleteDC lhDCC
        pDestroy
    End If
End Sub

Public Sub pDraw( _
       ByVal hdc As Long, _
       Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
       Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
       Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
                                        )
    If WidthSrc <= 0 Then WidthSrc = m_lWidth
    If HeightSrc <= 0 Then HeightSrc = m_lHeight
    BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy
End Sub

Private Sub pDestroy()
    If m_hBmpOld Then
        SelectObject m_ThDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If m_hBmp Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If m_ThDC Then
        DeleteDC m_ThDC
        m_ThDC = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
End Sub

Private Sub TXT_Change(Index As Integer)
    Select Case Index
        Case 0:
            Text = TXT(0).Text
            RaiseEvent Change
        Case 1:
            Text = TXT(1).Text
            RaiseEvent Change
        Case 2:
            Text = TXT(2).Text
            RaiseEvent Change
        Case 3:
            Text = TXT(3).Text
            RaiseEvent Change
    End Select
End Sub

Private Sub TXT_Click(Index As Integer)
    Select Case Index
        Case 0: RaiseEvent Click
        Case 1: RaiseEvent Click
        Case 2: RaiseEvent Click
        Case 3: RaiseEvent Click
    End Select
End Sub

Private Sub TXT_DblClick(Index As Integer)
    Select Case Index
        Case 0: RaiseEvent DblClick
        Case 1: RaiseEvent DblClick
        Case 2: RaiseEvent DblClick
        Case 3: RaiseEvent DblClick
    End Select
End Sub

Private Sub TXT_GotFocus(Index As Integer)
    Select Case Index
        Case 0:
            If xFocus = True Then
                CreateTextFocus TXT(0)
            End If
        Case 1:
            If xFocus = True Then
                CreateTextFocus TXT(1)
            End If
        Case 2:
            If xFocus = True Then
                CreateTextFocus TXT(2)
            End If
        Case 3:
            If xFocus = True Then
                CreateTextFocus TXT(3)
            End If
    End Select
End Sub

Private Sub TXT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0: RaiseEvent KeyDown(KeyCode, Shift)
        Case 1: RaiseEvent KeyDown(KeyCode, Shift)
        Case 2: RaiseEvent KeyDown(KeyCode, Shift)
        Case 3: RaiseEvent KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0: RaiseEvent KeyPress(KeyAscii)
        Case 1: RaiseEvent KeyPress(KeyAscii)
        Case 2: RaiseEvent KeyPress(KeyAscii)
        Case 3: RaiseEvent KeyPress(KeyAscii)
    End Select
End Sub

Private Sub TXT_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0: RaiseEvent KeyUp(KeyCode, Shift)
        Case 1: RaiseEvent KeyUp(KeyCode, Shift)
        Case 2: RaiseEvent KeyUp(KeyCode, Shift)
        Case 3: RaiseEvent KeyUp(KeyCode, Shift)
    End Select
End Sub

Private Sub TXT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: RaiseEvent MouseDown(Button, Shift, X, Y)
        Case 1: RaiseEvent MouseDown(Button, Shift, X, Y)
        Case 2: RaiseEvent MouseDown(Button, Shift, X, Y)
        Case 3: RaiseEvent MouseDown(Button, Shift, X, Y)
    End Select
End Sub

Private Sub TXT_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: RaiseEvent MouseMove(Button, Shift, X, Y)
        Case 1: RaiseEvent MouseMove(Button, Shift, X, Y)
        Case 2: RaiseEvent MouseMove(Button, Shift, X, Y)
        Case 3: RaiseEvent MouseMove(Button, Shift, X, Y)
    End Select
End Sub

Private Sub TXT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0: RaiseEvent MouseUp(Button, Shift, X, Y)
        Case 1: RaiseEvent MouseUp(Button, Shift, X, Y)
        Case 2: RaiseEvent MouseUp(Button, Shift, X, Y)
        Case 3: RaiseEvent MouseUp(Button, Shift, X, Y)
    End Select
End Sub

Private Sub UserControl_Initialize()
    hdc = UserControl.hdc
    m_Color = GetLngColor(vbHighlight)
    DoTextBoxStyler
End Sub

Private Sub UserControl_InitProperties()
    xBackColor = &HFFFFFF
    xForeColor = &H80000008
    xAlign = vbLeftJustify
    xLocked = False
    xPassword = ""
    xMax = 0
    xTExt = Ambient.DisplayName
    Set xFont = TXT(0).Font
    xScroll = vbSBNone
    xEnabled = True
    Set xIcon = Nothing
    xPointer = vbDefault
    xFocus = False
End Sub

Private Sub UserControl_Paint()
    DoTextBoxStyler
    If TXT(0).Visible = True Then
        RefreshAllObject 0
    ElseIf TXT(1).Visible = True Then
        RefreshAllObject 1
    ElseIf TXT(2).Visible = True Then
        RefreshAllObject 2
    ElseIf TXT(3).Visible = True Then
        RefreshAllObject 3
    End If
End Sub

Private Sub UserControl_Resize()
    Dim MyControlCount As Long
    On Error Resume Next
    hdc = UserControl.hdc
    For MyControlCount = 0 To TXT.Count - 1
        TXT(MyControlCount).Width = Picture1.Width - 30
        TXT(MyControlCount).Height = Picture2.Height - 30
        If MyControlCount > 0 Then
            TXT(MyControlCount).Top = TXT(0).Top
            TXT(MyControlCount).Left = TXT(0).Left
        End If
    Next
    If UserControl.Height <= 135 Then
        UserControl.Height = 255
    End If
End Sub

Private Sub UserControl_Show()
    If TXT(0).Visible = True Then
        RefreshAllObject 0
    ElseIf TXT(1).Visible = True Then
        RefreshAllObject 1
    ElseIf TXT(2).Visible = True Then
        RefreshAllObject 2
    ElseIf TXT(3).Visible = True Then
        RefreshAllObject 3
    End If
End Sub

Private Sub UserControl_Terminate()
    pDestroy
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property

Private Property Let hdc(ByVal cHdc As Long)
    m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
    If m_hDC = 0 Then
        m_hDC = UserControl.hdc
    Else
        m_MemDC = True
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = xBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    xBackColor = New_Color
    PropertyChanged "BackColor"
    If TXT(0).Visible = True Then
        RefreshBackColor 0
    ElseIf TXT(1).Visible = True Then
        RefreshBackColor 1
    ElseIf TXT(2).Visible = True Then
        RefreshBackColor 2
    ElseIf TXT(3).Visible = True Then
        RefreshBackColor 3
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = xForeColor
End Property

Public Property Let ForeColor(ByVal New_Color As OLE_COLOR)
    xForeColor = New_Color
    PropertyChanged "ForeColor"
    If TXT(0).Visible = True Then
        RefreshForeColor 0
    ElseIf TXT(1).Visible = True Then
        RefreshForeColor 1
    ElseIf TXT(2).Visible = True Then
        RefreshForeColor 2
    ElseIf TXT(3).Visible = True Then
        RefreshForeColor 3
    End If
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = xAlign
End Property

Public Property Let Alignment(ByVal New_Align As AlignmentConstants)
    xAlign = New_Align
    PropertyChanged "Alignment"
    If TXT(0).Visible = True Then
        RefreshAlign 0
    ElseIf TXT(1).Visible = True Then
        RefreshAlign 1
    ElseIf TXT(2).Visible = True Then
        RefreshAlign 2
    ElseIf TXT(3).Visible = True Then
        RefreshAlign 3
    End If
End Property

Public Property Get Locked() As Boolean
    Locked = xLocked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    xLocked = New_Locked
    PropertyChanged "Locked"
    If TXT(0).Visible = True Then
        RefreshLocked 0
    ElseIf TXT(1).Visible = True Then
        RefreshLocked 1
    ElseIf TXT(2).Visible = True Then
        RefreshLocked 2
    ElseIf TXT(3).Visible = True Then
        RefreshLocked 3
    End If
End Property

Public Property Get PasswordChar() As Variant
    PasswordChar = xPassword
End Property

Public Property Let PasswordChar(ByVal New_Pass As Variant)
    xPassword = New_Pass
    PropertyChanged "PasswordChar"
    If TXT(0).Visible = True Then
        RefreshPasswordChar 0
    ElseIf TXT(1).Visible = True Then
        RefreshPasswordChar 1
    ElseIf TXT(2).Visible = True Then
        RefreshPasswordChar 2
    ElseIf TXT(3).Visible = True Then
        RefreshPasswordChar 3
    End If
End Property

Public Property Get MaxLenght() As Long
    MaxLenght = xMax
End Property

Public Property Let MaxLenght(ByVal New_Max As Long)
    xMax = New_Max
    PropertyChanged "MaxLenght"
    If TXT(0).Visible = True Then
        RefreshMaxLenght 0
    ElseIf TXT(1).Visible = True Then
        RefreshMaxLenght 1
    ElseIf TXT(2).Visible = True Then
        RefreshMaxLenght 2
    ElseIf TXT(3).Visible = True Then
        RefreshMaxLenght 3
    End If
End Property

Public Property Get Text() As String
    Text = xTExt
End Property

Public Property Let Text(ByVal New_Text As String)
    xTExt = New_Text
    If TXT(0).Visible = True Then
        RefreshText 0
    ElseIf TXT(1).Visible = True Then
        RefreshText 1
    ElseIf TXT(2).Visible = True Then
        RefreshText 2
    ElseIf TXT(3).Visible = True Then
        RefreshText 3
    End If
    PropertyChanged "Text"
End Property

Public Property Get Font() As Font
    Set Font = xFont
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set xFont = New_Font
    PropertyChanged "Font"
    If TXT(0).Visible = True Then
        RefreshFont 0
    ElseIf TXT(1).Visible = True Then
        RefreshFont 1
    ElseIf TXT(2).Visible = True Then
        RefreshFont 2
    ElseIf TXT(3).Visible = True Then
        RefreshFont 3
    End If
End Property

Public Property Get Scrollbar() As ScrollBarConstants
    Scrollbar = xScroll
End Property

Public Property Let Scrollbar(ByVal New_Scroll As ScrollBarConstants)
    xScroll = New_Scroll
    PropertyChanged "Scrollbar"
    RefreshScroll
End Property

Public Property Get Enabled() As Boolean
    Enabled = xEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    xEnabled = New_Enabled
    PropertyChanged "Enabled"
    If TXT(0).Visible = True Then
        RefreshEnabled 0
    ElseIf TXT(1).Visible = True Then
        RefreshEnabled 1
    ElseIf TXT(2).Visible = True Then
        RefreshEnabled 2
    ElseIf TXT(3).Visible = True Then
        RefreshEnabled 3
    End If
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = xPointer
End Property

Public Property Let MousePointer(ByVal New_Pointer As MousePointerConstants)
    xPointer = New_Pointer
    PropertyChanged "MousePointer"
    If TXT(0).Visible = True Then
        RefreshPointer 0
    ElseIf TXT(1).Visible = True Then
        RefreshPointer 1
    ElseIf TXT(2).Visible = True Then
        RefreshPointer 2
    ElseIf TXT(3).Visible = True Then
        RefreshPointer 3
    End If
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = xIcon
End Property

Public Property Set MouseIcon(ByVal New_Icon As Picture)
    Set xIcon = New_Icon
    PropertyChanged "MouseIcon"
    If TXT(0).Visible = True Then
        RefreshMouseIcon 0
    ElseIf TXT(1).Visible = True Then
        RefreshMouseIcon 1
    ElseIf TXT(2).Visible = True Then
        RefreshMouseIcon 2
    ElseIf TXT(3).Visible = True Then
        RefreshMouseIcon 3
    End If
End Property

Public Property Get SelLength() As Long
    If TXT(0).Visible = True Then
        SelLength = TXT(0).SelLength
    ElseIf TXT(1).Visible = True Then
        SelLength = TXT(1).SelLength
    ElseIf TXT(2).Visible = True Then
        SelLength = TXT(2).SelLength
    ElseIf TXT(3).Visible = True Then
        SelLength = TXT(3).SelLength
    End If
End Property

Public Property Get SelText() As String
    If TXT(0).Visible = True Then
        SelText = TXT(0).SelText
    ElseIf TXT(1).Visible = True Then
        SelText = TXT(1).SelText
    ElseIf TXT(2).Visible = True Then
        SelText = TXT(2).SelText
    ElseIf TXT(3).Visible = True Then
        SelText = TXT(3).SelText
    End If
End Property

Public Property Get Focus() As Boolean
    Focus = xFocus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
    xFocus = New_Focus
    PropertyChanged "Focus"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", xBackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", xForeColor, &H80000008)
    Call PropBag.WriteProperty("Alignment", xAlign, vbLeftJustify)
    Call PropBag.WriteProperty("Locked", xLocked, False)
    Call PropBag.WriteProperty("PasswordChar", xPassword, "")
    Call PropBag.WriteProperty("MaxLenght", xMax, 0)
    Call PropBag.WriteProperty("Text", xTExt, "")
    Call PropBag.WriteProperty("Font", xFont, TXT(0).Font)
    Call PropBag.WriteProperty("Scrollbar", xScroll, vbSBNone)
    Call PropBag.WriteProperty("Enabled", xEnabled, True)
    Call PropBag.WriteProperty("MouseIcon", xIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", xPointer, vbDefault)
    Call PropBag.WriteProperty("Focus", xFocus, False)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    xBackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    xForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    xAlign = PropBag.ReadProperty("Alignment", vbLeftJustify)
    xLocked = PropBag.ReadProperty("Locked", False)
    xPassword = PropBag.ReadProperty("PasswordChar", "")
    xMax = PropBag.ReadProperty("MaxLenght", 0)
    xTExt = PropBag.ReadProperty("Text", "")
    Set xFont = PropBag.ReadProperty("Font", TXT(0).Font)
    xScroll = PropBag.ReadProperty("Scrollbar", vbSBNone)
    xEnabled = PropBag.ReadProperty("Enabled", True)
    Set xIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    xPointer = PropBag.ReadProperty("MousePointer", vbDefault)
    xFocus = PropBag.ReadProperty("Focus", False)
End Sub

Private Function RefreshPointer(Optional Index As Integer)
    TXT(Index).MousePointer = xPointer
End Function

Private Function RefreshMouseIcon(Optional Index As Integer)
    Set TXT(Index).MouseIcon = xIcon
End Function

Private Function RefreshEnabled(Optional Index As Integer)
    TXT(Index).Enabled = xEnabled
End Function

Private Function RefreshPasswordChar(Optional Index As Integer)
    TXT(Index).PasswordChar = xPassword
End Function

Private Function RefreshText(Optional Index As Integer)
    TXT(Index).Text = xTExt
End Function

Private Function RefreshLocked(Optional Index As Integer)
    TXT(Index).Locked = xLocked
End Function

Private Function RefreshAlign(Optional Index As Integer)
    TXT(Index).Alignment = xAlign
End Function

Private Function RefreshForeColor(Optional Index As Integer)
    TXT(Index).ForeColor = xForeColor
End Function

Private Function RefreshBackColor(Optional Index As Integer)
    TXT(Index).BackColor = xBackColor
End Function

Private Function RefreshMaxLenght(Optional Index As Integer)
    If xMax < 0 Then: xMax = 0: Exit Function
    TXT(Index).MaxLength = xMax
End Function

Private Function RefreshFont(Optional Index As Integer)
    Set TXT(Index).Font = xFont
End Function

Private Function RefreshScroll()
    If xScroll = vbSBNone Then
        TXT(0).Visible = True
        TXT(1).Visible = False
        TXT(2).Visible = False
        TXT(3).Visible = False

        RefreshPasswordChar 0
        RefreshText 0
        RefreshLocked 0
        RefreshAlign 0
        RefreshForeColor 0
        RefreshBackColor 0
        RefreshMaxLenght 0
        RefreshFont 0
        RefreshEnabled 0
    ElseIf xScroll = vbHorizontal Then
        TXT(0).Visible = False
        TXT(1).Visible = True
        TXT(2).Visible = False
        TXT(3).Visible = False

        RefreshPasswordChar 1
        RefreshText 1
        RefreshLocked 1
        RefreshAlign 1
        RefreshForeColor 1
        RefreshBackColor 1
        RefreshMaxLenght 1
        RefreshFont 1
        RefreshEnabled 1
    ElseIf xScroll = vbVertical Then
        TXT(0).Visible = False
        TXT(1).Visible = False
        TXT(2).Visible = True
        TXT(3).Visible = False

        RefreshPasswordChar 2
        RefreshText 2
        RefreshLocked 2
        RefreshAlign 2
        RefreshForeColor 2
        RefreshBackColor 2
        RefreshMaxLenght 2
        RefreshFont 2
        RefreshEnabled 2
    ElseIf xScroll = vbBoth Then
        TXT(0).Visible = False
        TXT(1).Visible = False
        TXT(2).Visible = False
        TXT(3).Visible = True

        RefreshPasswordChar 3
        RefreshText 3
        RefreshLocked 3
        RefreshAlign 3
        RefreshForeColor 3
        RefreshBackColor 3
        RefreshMaxLenght 3
        RefreshFont 3
        RefreshEnabled 3
    End If
End Function

Private Function RefreshAllObject(Optional Index As Integer)
    RefreshScroll
End Function

Private Function CreateTextFocus(ByVal ObjText As TextBox)
    With ObjText
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Function
