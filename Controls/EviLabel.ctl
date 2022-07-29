VERSION 5.00
Begin VB.UserControl EviLabel 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "EviLabel.ctx":0000
   ScaleHeight     =   675
   ScaleWidth      =   4800
   ToolboxBitmap   =   "EviLabel.ctx":003E
End
Attribute VB_Name = "EviLabel"
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
'enum for border style
Enum EviBorderStyleEnum
    eviNone = 0
    eviFixedSingle = 1
End Enum
Dim m_Appearance As EviAppearanceLabelEnum
Dim m_ForeColor As OLE_COLOR
Dim m_ShadowColor1 As OLE_COLOR
Dim m_ShadowColor2 As OLE_COLOR
Dim m_Caption As String
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawCaptionEviLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawCaptionEviLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawCaptionEviLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As EviBorderStyleEnum
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As EviBorderStyleEnum)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    DrawCaptionEviLabel
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawCaptionEviLabel
End Property

Public Property Get ShadowColor1() As OLE_COLOR
    ShadowColor1 = m_ShadowColor1
End Property

Public Property Let ShadowColor1(ByVal New_ShadowColor1 As OLE_COLOR)
    m_ShadowColor1 = New_ShadowColor1
    PropertyChanged "ShadowColor1"
    DrawCaptionEviLabel
End Property

Public Property Get ShadowColor2() As OLE_COLOR
    ShadowColor2 = m_ShadowColor2
End Property

Public Property Let ShadowColor2(ByVal New_ShadowColor2 As OLE_COLOR)
    m_ShadowColor2 = New_ShadowColor2
    PropertyChanged "ShadowColor2"
    DrawCaptionEviLabel
End Property

Public Property Get Appearance() As EviAppearanceLabelEnum
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As EviAppearanceLabelEnum)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    DrawCaptionEviLabel
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawCaptionEviLabel
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
    DrawCaptionEviLabel
End Sub

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

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error GoTo error
    Set UserControl.Font = Ambient.Font
    m_Appearance = eviNormalStyle
    m_ForeColor = vbBlack
    m_ShadowColor1 = &HC0C0C0
    m_ShadowColor2 = vbWhite
    m_Caption = Ambient.DisplayName
error:
End Sub

Private Sub UserControl_Paint()
    On Error GoTo error
    DrawCaptionEviLabel
error:
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    m_ShadowColor1 = PropBag.ReadProperty("ShadowColor1", &HC0C0C0)
    m_ShadowColor2 = PropBag.ReadProperty("ShadowColor2", vbWhite)
    m_Appearance = PropBag.ReadProperty("Appearance", 0)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
    On Error GoTo error
    DrawCaptionEviLabel
error:
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, vbBlack)
    Call PropBag.WriteProperty("ShadowColor1", m_ShadowColor1, &HC0C0C0)
    Call PropBag.WriteProperty("ShadowColor2", m_ShadowColor2, vbWhite)
    Call PropBag.WriteProperty("Appearance", m_Appearance, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Private Function DrawCaptionEviLabel()
    Dim I As Integer
    On Error GoTo error
    'clear all drawing
    UserControl.Cls

    'setting font evi label
    SettingLabelFont

    'if label enabled=true
    If UserControl.Enabled = True Then

        'if using shadow caption
        If m_Appearance = eviShadowCaption Then
            For I = 0 To 10
                CurrentX = I * 5
                CurrentY = I * 5
                UserControl.ForeColor = m_ShadowColor1
                UserControl.Print m_Caption
            Next
            For I = 0 To 10
                CurrentX = I - 30
                CurrentY = I - 30
                UserControl.ForeColor = m_ShadowColor2
                UserControl.Print m_Caption
            Next
            For I = 0 To 10
                CurrentX = 1
                CurrentY = 1
                UserControl.CurrentY = 10
                UserControl.ForeColor = m_ForeColor
                UserControl.Print m_Caption
            Next

            'if normal caption
        ElseIf m_Appearance = eviNormalStyle Then
            For I = 0 To 10
                CurrentX = 1
                CurrentY = 1
                UserControl.ForeColor = m_ForeColor
                UserControl.Print m_Caption
            Next
        End If

        'if label enabled=false
    Else
        For I = 0 To 10
            CurrentX = 1
            CurrentY = 1
            'constant forecolor if enabled=false
            UserControl.ForeColor = &H80000010
            UserControl.Print m_Caption
        Next
    End If
error:
End Function

Private Function SettingLabelFont()
    On Error GoTo error
    UserControl.FontBold = Font.Bold
    UserControl.FontItalic = Font.Italic
    UserControl.FontName = Font.Name
    UserControl.FontSize = Font.Size
    UserControl.FontStrikethru = Font.Strikethrough
    UserControl.FontUnderline = Font.Underline
error:
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
