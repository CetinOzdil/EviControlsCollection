VERSION 5.00
Begin VB.UserControl EviMaskedBox 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   PropertyPages   =   "EviMaskedBox.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   3645
   ToolboxBitmap   =   "EviMaskedBox.ctx":0048
   Begin VB.TextBox TXT 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   20
      TabIndex        =   2
      Top             =   20
      Width           =   3480
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   240
      Left            =   3510
      ScaleHeight     =   180
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3645
   End
End
Attribute VB_Name = "EviMaskedBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private m_MemDC As Boolean
Private xBackColor As OLE_COLOR
Private xForeColor As OLE_COLOR
Private m_Color As OLE_COLOR
Private xIcon As Picture
Private xPointer As MousePointerConstants
Private xFont As Font
Private xEnabled As Boolean
Private xMask As String
Private xTExt As String
Private xFocus As Boolean
Private xAlign As AlignmentConstants

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

Private Sub TXT_Change()
    xTExt = Format(TXT.Text, "")
    RaiseEvent Change
End Sub

Private Sub TXT_Click()
    RaiseEvent Click
End Sub

Private Sub TXT_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TXT_GotFocus()
    TXT.Text = Format(xTExt, "")
    If xFocus = True Then
        FocusMaskedText TXT
    End If
End Sub

Private Sub TXT_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TXT_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Exit Sub
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TXT_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TXT_LostFocus()
    If Len(xMask) <= 0 Then
        TXT.Text = xTExt
    Else
        TXT.Text = Format(xTExt, xMask)
    End If
End Sub

Private Sub TXT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TXT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub TXT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    hdc = UserControl.hdc
    m_Color = GetLngColor(vbHighlight)
    DoTextBoxStyler
End Sub

Private Sub UserControl_InitProperties()
    xBackColor = &H80000005
    xForeColor = &H80000008
    Set xIcon = Nothing
    xPointer = vbDefault
    xEnabled = True
    Set xFont = TXT.Font
    xMask = ""
    xTExt = ""
    xFocus = False
    TXT.MaxLength = 0
    TXT.Locked = False
    xAlign = vbLeftJustify
End Sub

Private Sub UserControl_Paint()
    DoTextBoxStyler
    RefreshObjectMask
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    hdc = UserControl.hdc
    TXT.Width = Picture1.Width - 30
    TXT.Height = Picture2.Height - 30
    If UserControl.Height <= 135 Then
        UserControl.Height = 255
    End If
End Sub

Private Sub UserControl_Show()
    RefreshObjectMask
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
    RefreshObjectMask
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = xForeColor
End Property

Public Property Let ForeColor(ByVal New_Color As OLE_COLOR)
    xForeColor = New_Color
    PropertyChanged "ForeColor"
    RefreshObjectMask
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = xIcon
End Property

Public Property Set MouseIcon(ByVal New_Icon As Picture)
    Set xIcon = New_Icon
    PropertyChanged "MouseIcon"
    RefreshObjectMask
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = xPointer
End Property

Public Property Let MousePointer(ByVal New_Pointer As MousePointerConstants)
    xPointer = New_Pointer
    PropertyChanged "MousePointer"
    RefreshObjectMask
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = xEnabled
End Property

Public Property Let Enabled(ByVal New_E As Boolean)
    xEnabled = New_E
    PropertyChanged "Enabled"
    RefreshObjectMask
End Property

Public Property Get Font() As Font
    Set Font = xFont
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set xFont = New_Font
    PropertyChanged "Font"
    RefreshObjectMask
End Property

Public Property Get Mask() As String
Attribute Mask.VB_ProcData.VB_Invoke_Property = "General"
    Mask = xMask
End Property

Public Property Let Mask(ByVal New_Mask As String)
    xMask = New_Mask
    PropertyChanged "Mask"
    RefreshObjectMask
End Property

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "General"
    Text = xTExt
End Property

Public Property Get FormattedText() As String
    FormattedText = Format(xTExt, xMask)
End Property

Public Property Let Text(ByVal New_Text As String)
    xTExt = New_Text
    PropertyChanged "Text"
    RefreshObjectMask
End Property

Public Property Get Focus() As Boolean
Attribute Focus.VB_ProcData.VB_Invoke_Property = "General"
    Focus = xFocus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
    xFocus = New_Focus
    PropertyChanged "Focus"
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = xAlign
End Property

Public Property Let Alignment(ByVal New_Align As AlignmentConstants)
    xAlign = New_Align
    PropertyChanged "Alignment"
    RefreshObjectMask
End Property

Private Function RefreshObjectMask()
    TXT.BackColor = xBackColor
    TXT.ForeColor = xForeColor
    Set TXT.MouseIcon = xIcon
    TXT.MousePointer = xPointer
    TXT.Enabled = xEnabled
    Set TXT.Font = xFont
    TXT.Text = Format(xTExt, xMask)
    TXT.Alignment = xAlign
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    xBackColor = PropBag.ReadProperty("BackColor", &H80000005)
    xForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set xIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    xPointer = PropBag.ReadProperty("MousePointer", vbDefault)
    xEnabled = PropBag.ReadProperty("Enabled", True)
    Set xFont = PropBag.ReadProperty("Font", TXT.Font)
    xMask = PropBag.ReadProperty("Mask", "")
    xFocus = PropBag.ReadProperty("Focus", False)
    xTExt = PropBag.ReadProperty("Text", "")
    TXT.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    TXT.Locked = PropBag.ReadProperty("Locked", False)
    xAlign = PropBag.ReadProperty("Alignment", vbLeftJustify)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", xBackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", xForeColor, &H80000008)
    Call PropBag.WriteProperty("MouseIcon", xIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", xPointer, vbDefault)
    Call PropBag.WriteProperty("Enabled", xEnabled, True)
    Call PropBag.WriteProperty("Font", xFont, TXT.Font)
    Call PropBag.WriteProperty("Mask", xMask, "")
    Call PropBag.WriteProperty("Focus", xFocus, False)
    Call PropBag.WriteProperty("MaxLength", TXT.MaxLength, 0)
    Call PropBag.WriteProperty("Text", xTExt, "")
    Call PropBag.WriteProperty("Locked", TXT.Locked, False)
    Call PropBag.WriteProperty("Alignment", xAlign, vbLeftJustify)
End Sub

Private Function FocusMaskedText(ByVal ObjMask As Object)
    With ObjMask
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TXT,TXT,-1,MaxLength
Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Sets/returns the maximum length of the masked edit control."
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = "General"
    MaxLength = TXT.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Integer)
    TXT.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TXT,TXT,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Attribute Locked.VB_ProcData.VB_Invoke_Property = "General"
    Locked = TXT.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TXT.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
