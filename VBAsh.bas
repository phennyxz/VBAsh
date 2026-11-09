Option Explicit

Public gSession As clsSession
Public gConsole As clsConsole
Public gFDTable As clsFileDescriptorTable

Public Const VBAsh_STREAM_ERROR_NOT_SUPPORTED As Long = vbObjectError + 1200
Public Const VBAsh_STREAM_ERROR_CLOSED As Long = vbObjectError + 1201
Public Const VBAsh_STREAM_ERROR_READ_FAILURE As Long = vbObjectError + 1202
Public Const VBAsh_STREAM_ERROR_WRITE_FAILURE As Long = vbObjectError + 1203
Public Const VBAsh_STREAM_ERROR_OPEN_FAILURE As Long = vbObjectError + 1204

Public Sub Auto_Open()
    VBAsh_Start
End Sub

Public Sub VBAsh_Start()
    Dim terminal As Object
    Dim failure As String
    Dim standardInput As clsNullStream
    Dim standardOutput As clsMemoryStream
    Dim standardError As clsMemoryStream

    Set gSession = New clsSession
    Set gConsole = New clsConsole
    Set gFDTable = New clsFileDescriptorTable
    Set standardInput = New clsNullStream
    standardInput.Initialize True, False
    Set standardOutput = New clsMemoryStream
    standardOutput.Initialize "", True, True
    Set standardError = New clsMemoryStream
    standardError.Initialize "", True, True
    gFDTable.Initialize standardInput, standardOutput, standardError, True
    gSession.ConfigureDescriptorTable gFDTable, standardOutput, standardError

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