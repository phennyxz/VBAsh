Option Explicit

Public TokenType As String
Public Value As String
Public ExpansionMask As String
Public BraceExpansionAllowed As Boolean
Public Segments As Collection
' Token offsets are 1-based inclusive spans in the original Tokenize input.
Public SourceStart As Long
Public SourceEnd As Long

Private Sub Class_Initialize()
    Set Segments = New Collection
End Sub

Public Sub Init(ByVal pType As String, ByVal pValue As String, Optional ByVal pExpansionMask As String = "", _
                Optional ByVal pBraceExpansionAllowed As Boolean = True)
    TokenType = pType
    Value = pValue
    ExpansionMask = pExpansionMask
    BraceExpansionAllowed = pBraceExpansionAllowed
    Set Segments = New Collection
    SourceStart = 0
    SourceEnd = 0
End Sub