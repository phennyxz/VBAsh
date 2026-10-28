Option Explicit

Public TokenType As String
Public Value As String
Public ExpansionMask As String

Public Sub Init(ByVal pType As String, ByVal pValue As String, Optional ByVal pExpansionMask As String = "")
    TokenType = pType
    Value = pValue
    ExpansionMask = pExpansionMask
End Sub