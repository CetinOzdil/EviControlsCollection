Attribute VB_Name = "ModGlobal"
Option Explicit

Private Declare Function ComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Property Get GetComputerName() As String
    Dim sBuffer As String
    Dim lAns As Long

    On Error GoTo error

    sBuffer = Space$(255)
    lAns = ComputerName(sBuffer, 255)

    If lAns <> 0 Then
        GetComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If

error:
End Property

Public Sub DebugPrint(ByVal data As String)
    Debug.Print data
    Call OutputDebugString(data)
End Sub

