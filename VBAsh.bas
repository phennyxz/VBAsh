Option Explicit

Public gSession As clsSession
Public gConsole As clsConsole
Public gFDTable As clsFileDescriptorTable

Public Sub Auto_Open()
    VBAsh_Start
End Sub

Public Sub VBAsh_Start()
    Dim terminal As Object
    Dim failure As String

    Set gSession = New clsSession
    Set gConsole = New clsConsole
    Set gFDTable = New clsFileDescriptorTable

    gConsole.Initialize

    On Error GoTo RestoreWorkbook
    ThisWorkbook.Windows(1).Visible = False
    Set terminal = VBA.UserForms.Add("frmVBAsh")
    terminal.Show
    GoTo Finish

RestoreWorkbook:
    failure = Err.Description

Finish:
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = True
    On Error GoTo 0
    If Len(failure) > 0 Then MsgBox "Não foi possível abrir o terminal VBAsh:" & vbCrLf & failure, vbCritical, "VBAsh"
End Sub