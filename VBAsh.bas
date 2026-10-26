Option Explicit

Public gSession As clsSession
Public gConsole As clsConsole
Public gFDTable As clsFileDescriptorTable

Public Sub Auto_Open()
    VBAsh_Start
End Sub

Public Sub VBAsh_Start()
    Dim terminal As Object

    Set gSession = New clsSession
    Set gConsole = New clsConsole
    Set gFDTable = New clsFileDescriptorTable

    gConsole.Initialize

    On Error GoTo RestoreWorkbook
    ThisWorkbook.Windows(1).Visible = False
    Set terminal = VBA.UserForms.Add("frmVBAsh")
    terminal.Show

RestoreWorkbook:
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = True
    On Error GoTo 0
End Sub