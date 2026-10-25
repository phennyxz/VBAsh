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
    Set Args = New Collection
    Set Redirections = New Collection
End Sub

Public Sub AddArg(ByVal arg As String)
    Args.Add arg
End Sub

Public Sub AddRedirection(ByVal redirection As clsRedirection)
    Redirections.Add redirection
End Sub