Option Explicit

Public TokenType As String
Public Value As String

Public Sub Init(ByVal pType As String, ByVal pValue As String)
    TokenType = pType
    Value = pValue
End Sub