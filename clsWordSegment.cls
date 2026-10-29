Option Explicit

Public Text As String
' QuoteContext is UNQUOTED, SINGLE, or DOUBLE; escaped characters form their own segment.
Public QuoteContext As String
Public IsEscaped As Boolean
' Offsets are 1-based inclusive; non-empty segment spans cover source characters, not quote marks.
Public SourceStart As Long
Public SourceEnd As Long

Public Sub Init(ByVal pText As String, ByVal pQuoteContext As String, ByVal pIsEscaped As Boolean, _
                ByVal pSourceStart As Long, ByVal pSourceEnd As Long)
    Text = pText
    QuoteContext = pQuoteContext
    IsEscaped = pIsEscaped
    SourceStart = pSourceStart
    SourceEnd = pSourceEnd
End Sub
