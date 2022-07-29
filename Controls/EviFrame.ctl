VERSION 5.00
Begin VB.UserControl EviFrame 
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ControlContainer=   -1  'True
   DrawWidth       =   5
   PropertyPages   =   "EviFrame.ctx":0000
   ScaleHeight     =   1995
   ScaleWidth      =   2775
   ToolboxBitmap   =   "EviFrame.ctx":003E
   Begin VB.Image Image2 
      Height          =   255
      Left            =   600
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   100
      Stretch         =   -1  'True
      Top             =   90
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   0
      Top             =   420
      Width           =   2775
   End
End
Attribute VB_Name = "EviFrame"
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
Private m_Color As OLE_COLOR
'public enum gradient position
Enum GradientPositionEnum
    [Top Gradient] = &H0
    [Bottom Gradient] = &H1
    [Left Gradient] = &H2
    [Right Gradient] = &H3
    [Spin Gradient] = &H4
    [Box Gradient] = &H5
    [Diagonal Gradient] = &H6
    [Rectangular Gradient] = &H7
    [Circle Gradient] = &H8
End Enum
'public enum Style Frame
Enum StyleEviFrameEnum
    [XP] = 0
    [Office 2003] = 1
    [Longhorn] = 2
    [Vista] = 3
End Enum
'private Style Enum
Private m_Style As StyleEviFrameEnum
Private m_Theme As ThemeEviFrameEnum
Private m_GradientPosition As GradientPositionEnum
'public enum theme frame
Enum ThemeEviFrameEnum
    [Blue Theme] = 0
    [Olive Theme] = 1
    [Silver Theme] = 2
    [Royal Theme] = 3
    [Black Theme] = 4
    [Red Theme] = 5
End Enum
Enum EviAppearanceLabelEnum
    eviNormalStyle = 0
    eviShadowCaption = 1
End Enum
Private m_GradientColor1 As OLE_COLOR
Private m_GradientColor2 As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_Caption As String
Private m_ForeColor As OLE_COLOR
Private m_ShadowColor As OLE_COLOR
Private m_Appearance As EviAppearanceLabelEnum
Private m_ShowIcon As Boolean
Private m_ForeDisabled As OLE_COLOR
Private m_Icon As Picture
Private m_IconDisabled As Picture
'Event Declarations:
Event Click()        'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()        'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)        'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)        'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)        'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize()

Public Property Get ShowIcon() As Boolean
    ShowIcon = m_ShowIcon
End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)
    m_ShowIcon = New_ShowIcon
    PropertyChanged "ShowIcon"
    DrawingTitleEviFrame
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawingTitleEviFrame
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawingTitleEviFrame
End Property

Private Function DrawingTitleEviFrame()
    On Error GoTo error
    Shape1.BackColor = m_BackColor
    DrawGradientTitle m_GradientPosition, 25, UserControl.ScaleWidth
    Shape1.BorderColor = m_Color

    UserControl.FontBold = Font.Bold
    UserControl.FontItalic = Font.Italic
    UserControl.FontName = Font.Name
    UserControl.FontSize = Font.Size
    UserControl.FontStrikethru = Font.Strikethrough
    UserControl.FontUnderline = Font.Underline
    If UserControl.Enabled = True Then
        If m_ShowIcon = True Then
            Image1.Visible = True
            Image2.Visible = False
        Else
            Image1.Visible = False
            Image2.Visible = False
        End If
    Else
        If m_ShowIcon = True Then
            Image1.Visible = False
            Image2.Visible = True
        Else
            Image1.Visible = False
            Image2.Visible = False
        End If
    End If
    Image1.Picture = m_Icon
    Image2.Picture = m_IconDisabled
    DrawCaptionEviLabel
error:
End Function

Private Sub DrawGradientTitle(Optional NewGradient As GradientPositionEnum, Optional SH As Long, Optional SW As Long)
    Dim VR, VG, VB As Single
    Dim Color1, Color2 As OLE_COLOR
    Dim R, G, b, R2, G2, X, Y, B2 As Integer
    Dim temp As Long
    Dim m_Position, m_Right, m_Left As Long
    On Error GoTo error

    UserControl.Cls

    UserControl.AutoRedraw = True
    'UserControl.DrawWidth = 5
    UserControl.ScaleMode = vbPixels

    Color1 = m_GradientColor1
    Color2 = m_GradientColor2

    m_Position = 0
    m_Right = 0
    m_Left = 0

    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    If NewGradient = [Top Gradient] Then

        VR = Abs(R - R2) / SH
        VG = Abs(G - G2) / SH
        VB = Abs(b - B2) / SH

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For Y = 0 To SH
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = b + VB * Y
            UserControl.Line (0, Y)-(SW, Y), RGB(R2, G2, B2)
        Next Y

    ElseIf NewGradient = [Bottom Gradient] Then
        m_Position = SH / 30
        m_Left = SH - m_Position
        m_Right = m_Left + m_Position

        VR = Abs(R - R2) / SH
        VG = Abs(G - G2) / SH
        VB = Abs(b - B2) / SH

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For Y = 0 To SH
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = b + VB * Y

            UserControl.Line (0, m_Left)-(SW, m_Right), RGB(R2, G2, B2), BF
            m_Left = m_Left - m_Position
            m_Right = m_Left + m_Position
        Next Y

    ElseIf NewGradient = [Left Gradient] Then

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = 0 To SW
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line (X, 0)-(X, SH), RGB(R2, G2, B2)
        Next X

    ElseIf NewGradient = [Right Gradient] Then

        m_Position = SW / 200
        m_Left = SW - m_Position
        m_Right = m_Left + m_Position

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = 0 To SW
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line (m_Left, 0)-(m_Right, SH), RGB(R2, G2, B2)

            m_Left = m_Left - m_Position
            m_Right = m_Left + m_Position

        Next X

    ElseIf NewGradient = [Spin Gradient] Then

        VR = Abs(R - R2) / SW / 2
        VG = Abs(G - G2) / SW / 2
        VB = Abs(b - B2) / SW / 2

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = 0 To SW
            R2 = R2 - VR        '* X
            G2 = G2 - VG        '* X
            B2 = B2 - VB        '* X
            UserControl.Line (X, 0)-(SW - X, SH), RGB(R2, G2, B2)
        Next X

        For X = 0 To SH
            R2 = R2 - VR        '* X
            G2 = G2 - VG        '* X
            B2 = B2 - VB        '* X
            UserControl.Line (SW, X)-(0, SH - X), RGB(R2, G2, B2)
        Next X

    ElseIf NewGradient = [Box Gradient] Then

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = SW To 0 Step -1
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line ((X / 5), (X / 5))-(SW - (X / 5), SH - (X / 5)), RGB(R2, G2, B2), B

        Next X

    ElseIf NewGradient = [Diagonal Gradient] Then

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = 0 To SW
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line (0, X)-(X, 0), RGB(R2, G2, B2)
        Next X

        For X = SW To 0 Step -1
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line (SW - X, SW)-(SW, SW - X), RGB(R2, G2, B2)
        Next X

    ElseIf NewGradient = [Rectangular Gradient] Then

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = SW To 0 Step -1
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Line ((X / 2), (X / 2))-(SW - (X / 2), SH - (X / 2)), RGB(R2, G2, B2), B
        Next X

    ElseIf NewGradient = [Circle Gradient] Then

        VR = Abs(R - R2) / SW
        VG = Abs(G - G2) / SW
        VB = Abs(b - B2) / SW

        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB

        For X = 0 To SW
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X

            UserControl.Circle (SW / 2, SH / 2), X, RGB(R2, G2, B2)
        Next X
    End If
error:
End Sub

Private Sub UserControl_Initialize()
    Image2.Top = Image1.Top
    Image2.Left = Image1.Left
    Image2.Height = Image1.Height
    Image2.Width = Image1.Width
End Sub

Private Sub UserControl_InitProperties()
    On Error GoTo error
    Set UserControl.Font = Ambient.Font
    m_GradientColor1 = vbWhite
    m_GradientColor2 = &HF7D3C6
    m_Color = &HF7D3C6
    m_BackColor = &HF7D3C6
    m_GradientPosition = [Left Gradient]
    m_Caption = Ambient.DisplayName
    m_ForeColor = Ambient.ForeColor
    m_ShadowColor = &HC0C0C0
    m_Appearance = eviNormalStyle
    m_ShowIcon = False
    m_ForeDisabled = &H80000010
    Set m_Icon = Nothing
    Set m_IconDisabled = Nothing
    DrawingTitleEviFrame
error:
End Sub

Private Sub UserControl_Paint()
    On Error GoTo error
    DrawingTitleEviFrame
error:
End Sub

Private Sub UserControl_Resize()
    On Error GoTo error
    DrawingTitleEviFrame
    Shape1.Width = UserControl.ScaleWidth
    Shape1.Height = UserControl.ScaleHeight - 28
    RaiseEvent Resize
error:
End Sub

Private Sub UserControl_Show()
    On Error GoTo error
    DrawingTitleEviFrame
error:
End Sub

Public Property Get GradientColor1() As OLE_COLOR
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"
    DrawingTitleEviFrame
End Property

Public Property Get GradientColor2() As OLE_COLOR
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"
    DrawingTitleEviFrame
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_Color
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_Color = New_BorderColor
    PropertyChanged "BorderColor"
    DrawingTitleEviFrame
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawingTitleEviFrame
End Property

Public Property Get GradientPosition() As GradientPositionEnum
    GradientPosition = m_GradientPosition
End Property

Public Property Let GradientPosition(ByVal New_GradientPosition As GradientPositionEnum)
    m_GradientPosition = New_GradientPosition
    PropertyChanged "GradientPosition"
    DrawingTitleEviFrame
End Property

Public Property Get Theme() As ThemeEviFrameEnum
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As ThemeEviFrameEnum)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    Select Case Theme
        Case 0:
            If m_Style = XP Then
                DrawingFrameWithGradient vbWhite, &HF7D3C6, &HF7DFD6 _
                                                            , &HF7DFD6, [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient vbWhite, &HF7D3C6, &HF7DFD6 _
                                                            , &HF7DFD6, [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, &HF7D3C6, vbWhite _
                                                            , vbWhite, [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, &HF7D3C6, vbWhite _
                                                            , vbWhite, [Circle Gradient]
            End If
        Case 1:
            If m_Style = XP Then
                DrawingFrameWithGradient vbWhite, &HB8E7E0, &HECF6F6 _
                                                            , &HECF6F6, [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient vbWhite, &HB8E7E0, &HECF6F6 _
                                                            , &HECF6F6, [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, &HB8E7E0, vbWhite _
                                                            , vbWhite, [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, &HB8E7E0, vbWhite _
                                                            , vbWhite, [Circle Gradient]
            End If
        Case 2:
            If m_Style = XP Then
                DrawingFrameWithGradient vbWhite, &HE0D7D6, &HF5F1F0 _
                                                            , &HF5F1F0, [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient vbWhite, &HE0D7D6, &HF5F1F0 _
                                                            , &HF5F1F0, [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, &HE0D7D6, vbWhite _
                                                            , vbWhite, [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, &HE0D7D6, vbWhite _
                                                            , vbWhite, [Circle Gradient]
            End If
        Case 3:
            If m_Style = XP Then
                DrawingFrameWithGradient vbBlue, &H808000, &HB75F31, _
                                         &HFFFFFF, [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient vbBlue, &H808000, &HB75F31, _
                                         &HFFFFFF, [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, vbBlue, vbWhite, _
                                         vbWhite, [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, vbBlue, vbWhite, _
                                         vbWhite, [Circle Gradient]
            End If
        Case 4:
            If m_Style = XP Then
                DrawingFrameWithGradient &HC0C0C0, vbBlack, vbBlack, &HE0E0E0 _
                                                                     , [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient &HC0C0C0, vbBlack, vbBlack, &HE0E0E0 _
                                                                     , [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, vbBlack, vbWhite, vbWhite _
                                                                    , [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, vbBlack, vbWhite, vbWhite _
                                                                    , [Circle Gradient]
            End If
        Case 5:
            If m_Style = XP Then
                DrawingFrameWithGradient &H8080FF, &H80&, &H80&, _
                                         &HC0C0FF, [Left Gradient]
            ElseIf m_Style = [Office 2003] Then
                DrawingFrameWithGradient &H8080FF, &H80&, &H80&, _
                                         &HC0C0FF, [Top Gradient]
            ElseIf m_Style = Longhorn Then
                DrawingFrameWithGradient vbWhite, &H80&, vbWhite, _
                                         vbWhite, [Bottom Gradient]
            ElseIf m_Style = Vista Then
                DrawingFrameWithGradient vbWhite, &H80&, vbWhite, _
                                         vbWhite, [Circle Gradient]
            End If
    End Select
End Property

Public Property Get Style() As StyleEviFrameEnum
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As StyleEviFrameEnum)
    m_Style = New_Style
    PropertyChanged "Style"
    Select Case m_Style
        Case 0: m_GradientPosition = [Left Gradient]
        Case 1: m_GradientPosition = [Top Gradient]
        Case 2: m_GradientPosition = [Bottom Gradient]
    End Select
    DrawingTitleEviFrame
    Theme = m_Theme
End Property

Private Sub DrawingFrameWithGradient(Optional GC1 As OLE_COLOR, Optional GC2 As OLE_COLOR _
                                                                , Optional BrC As OLE_COLOR, Optional BkC As OLE_COLOR, Optional _
                                                                                                                        GP As GradientPositionEnum, Optional CC As OLE_COLOR, Optional _
                                                                                                                                                                              CB As OLE_COLOR, Optional CT As OLE_COLOR)
    On Error GoTo error
    m_BackColor = BkC
    m_Color = BrC
    m_GradientColor1 = GC1
    m_GradientColor2 = GC2
    Shape1.BorderColor = m_Color
    Shape1.BackColor = m_BackColor
    m_GradientPosition = GP
    DrawingTitleEviFrame
error:
End Sub

Private Function DrawCaptionEviLabel()
    Dim I As Integer
    On Error GoTo error
    'if label enabled=true
    If UserControl.Enabled = True Then

        'if using shadow caption
        If m_Appearance = eviShadowCaption Then
            For I = 0 To 10
                If m_ShowIcon = False Then
                    CurrentX = 2.6 * 2
                Else
                    If m_Icon Is Nothing Then
                        CurrentX = 2.6 * 2
                    Else
                        CurrentX = 12.6 * 2
                    End If
                End If
                CurrentY = 3.7 * 2
                UserControl.ForeColor = m_ShadowColor
                UserControl.Print m_Caption
            Next
            For I = 0 To 10
                If m_ShowIcon = False Then
                    CurrentX = 6
                Else
                    If m_Icon Is Nothing Then
                        CurrentX = 6
                    Else
                        CurrentX = 26
                    End If
                End If
                CurrentY = 6
                UserControl.ForeColor = m_ForeColor
                UserControl.Print m_Caption
            Next

            'if normal caption
        ElseIf m_Appearance = eviNormalStyle Then
            For I = 0 To 10
                If m_ShowIcon = False Then
                    CurrentX = 6
                Else
                    If m_Icon Is Nothing Then
                        CurrentX = 6
                    Else
                        CurrentX = 26
                    End If
                End If
                CurrentY = 6
                UserControl.ForeColor = m_ForeColor
                UserControl.Print m_Caption
            Next
        End If

        'if label enabled=false
    Else
        For I = 0 To 10
            If m_ShowIcon = False Then
                CurrentX = 6
            Else
                If m_IconDisabled Is Nothing Then
                    CurrentX = 6
                Else
                    CurrentX = 26
                End If
            End If
            CurrentY = 6
            'constant forecolor if enabled=false
            UserControl.ForeColor = m_ForeDisabled
            UserControl.Print m_Caption
        Next
    End If
error:
End Function

Public Property Get ForeDisabled() As OLE_COLOR
    ForeDisabled = m_ForeDisabled
End Property

Public Property Let ForeDisabled(ByVal New_ForeDisabled As OLE_COLOR)
    m_ForeDisabled = New_ForeDisabled
    PropertyChanged "ForeDisabled"
    DrawingTitleEviFrame
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    m_ShadowColor = New_ShadowColor
    PropertyChanged "ShadowColor"
    DrawingTitleEviFrame
End Property

Public Property Get Appearance() As EviAppearanceLabelEnum
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As EviAppearanceLabelEnum)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    DrawingTitleEviFrame
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_GradientColor1 = PropBag.ReadProperty("GradientColor1", vbWhite)
    m_GradientColor2 = PropBag.ReadProperty("GradientColor2", &HF7D3C6)
    m_Color = PropBag.ReadProperty("BorderColor", &HF7D3C6)
    m_BackColor = PropBag.ReadProperty("BackColor", &HF7D3C6)
    m_GradientPosition = PropBag.ReadProperty("GradientPosition", &H2)
    m_Theme = PropBag.ReadProperty("Theme", 0)
    m_Style = PropBag.ReadProperty("Style", 0)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", &HC0C0C0)
    m_Appearance = PropBag.ReadProperty("Appearance", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ForeDisabled = PropBag.ReadProperty("ForeDisabled", &H80000010)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    Set m_IconDisabled = PropBag.ReadProperty("IconDisabled", Nothing)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ShowIcon = PropBag.ReadProperty("ShowIcon", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GradientColor1", m_GradientColor1, vbWhite)
    Call PropBag.WriteProperty("GradientColor2", m_GradientColor2, &HF7D3C6)
    Call PropBag.WriteProperty("BorderColor", m_Color, &HF7D3C6)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HF7D3C6)
    Call PropBag.WriteProperty("GradientPosition", m_GradientPosition, &H2)
    Call PropBag.WriteProperty("Theme", m_Theme, 0)
    Call PropBag.WriteProperty("Style", m_Style, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("ShadowColor", m_ShadowColor, &HC0C0C0)
    Call PropBag.WriteProperty("Appearance", m_Appearance, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeDisabled", m_ForeDisabled, &H80000010)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("IconDisabled", m_IconDisabled, Nothing)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ShowIcon", m_ShowIcon, False)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawingTitleEviFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
    DrawingTitleEviFrame
End Property

Public Property Get IconDisabled() As Picture
    Set IconDisabled = m_IconDisabled
End Property

Public Property Set IconDisabled(ByVal New_IconDisabled As Picture)
    Set m_IconDisabled = New_IconDisabled
    PropertyChanged "IconDisabled"
    DrawingTitleEviFrame
End Property
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawingTitleEviFrame
End Property
