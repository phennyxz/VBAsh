Option Explicit

' LastErrorNumber and LastErrorDescription describe the most recent failed operation.
' A successful read, write, or close clears that error state.
' Unsupported operations and operations on a closed stream raise
' VBAsh_STREAM_ERROR_NOT_SUPPORTED and update the error state.
' AtEndOfStream is true when no more data can be read or the stream is closed.
' ReadText returns at most maxChars; a non-positive limit returns empty text.
' WriteText returns the number of characters accepted. CloseStream is idempotent.
Public Property Get CanRead() As Boolean
End Property

Public Property Get CanWrite() As Boolean
End Property

Public Property Get AtEndOfStream() As Boolean
End Property

Public Property Get IsClosed() As Boolean
End Property

Public Property Get HasError() As Boolean
End Property

Public Property Get LastErrorNumber() As Long
End Property

Public Property Get LastErrorDescription() As String
End Property

Public Function ReadText(ByVal maxChars As Long) As String
End Function

Public Function WriteText(ByVal text As String) As Long
End Function

Public Sub CloseStream()
End Sub
