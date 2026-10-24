Option Explicit

Public CurrentDirectory As String
Public Env As Object
Public History As Collection
Public LastExitCode As Long
Public Prompt As String

Private Sub Class_Initialize()
    CurrentDirectory = CurDir$
    Set Env = CreateObject("Scripting.Dictionary")
    Set History = New Collection
    LastExitCode = 0
    Prompt = "vbash> "
End Sub

Public Sub SetEnv(ByVal key As String, ByVal value As String)
    Env(key) = value
End Sub

Public Function GetEnv(ByVal key As String) As String
    If Env.Exists(key) Then
        GetEnv = Env(key)
    Else
        GetEnv = ""
    End If
End Function

Public Sub AddHistory(ByVal commandLine As String)
    If Len(Trim$(commandLine)) > 0 Then
        History.Add commandLine
    End If
End Sub

Public Function ProcessInput(ByVal commandLine As String) As String
    Dim trimmed As String
    trimmed = Trim$(commandLine)

    If Len(trimmed) = 0 Then
        ProcessInput = ""
        Exit Function
    End If

    AddHistory trimmed

    Select Case LCase$(trimmed)
        Case "pwd"
            ProcessInput = CurrentDirectory

        Case "exit"
            ProcessInput = "exit"

        Case "help"
            ProcessInput = "VBAsh minimal shell. Builtins available: pwd, help, exit, setenv."

        Case "echo"
            ProcessInput = "echo: command runner not implemented yet."

        Case Else
            If InStr(1, trimmed, "=") > 0 Then
                Dim parts() As String
                parts = Split(trimmed, "=", 2)
                SetEnv parts(0), parts(1)
                ProcessInput = ""
            Else
                ProcessInput = "Command not recognized: " & trimmed
                LastExitCode = 127
            End If
    End Select
End Function