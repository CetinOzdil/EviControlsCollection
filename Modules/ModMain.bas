Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub InitCommonControlsVB()
    IsUserAnAdmin
    InitCommonControls
End Sub

Public Sub Main()
    InitCommonControlsVB
    Sample.Show
End Sub


