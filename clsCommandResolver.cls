Option Explicit

Public Function Resolve(ByVal commandName As String) As String
    Select Case LCase$(commandName)
        Case "echo", "pwd", "cd", "export", "setenv", "unset", "help", "exit"
            Resolve = LCase$(commandName)
        Case Else
            Resolve = ""
    End Select
End Function