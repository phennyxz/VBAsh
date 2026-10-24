Option Explicit

Public Name As String
Public Args As Collection
Public Redirections As Collection
Public OriginalText As String

Private Sub Class_Initialize()
    Set Args = New Collection
    Set Redirections = New Collection
End Sub

Public Sub Init(ByVal pName As String, ByVal pOriginalText As String)
    Name = pName
    OriginalText = pOriginalText
    Args.RemoveAll
    Redirections.RemoveAll
End Sub

Public Sub AddArg(ByVal arg As String)
    If Len(arg) > 0 Then
        Args.Add arg
    End If
End Sub

Public Sub AddRedirection(ByVal redirection As String)
    Redirections.Add redirection
End Sub