Option Explicit

Public TokenType As String
Public Value As String
Public ExpansionMask As String
Public BraceExpansionAllowed As Boolean

Public Sub Init(ByVal pType As String, ByVal pValue As String, Optional ByVal pExpansionMask As String = "", _
                Optional ByVal pBraceExpansionAllowed As Boolean = True)
    TokenType = pType
    Value = pValue
    ExpansionMask = pExpansionMask
    BraceExpansionAllowed = pBraceExpansionAllowed
End Sub