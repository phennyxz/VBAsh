VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    If Len(arg) > 0 Then
        Args.Add arg
    End If
End Sub

Public Sub AddRedirection(ByVal redirection As String)
    Redirections.Add redirection
End Sub
