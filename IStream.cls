Option Explicit

' LastErrorNumber and LastErrorDescription describe the most recent failed operation.
' A successful read, write, or close clears that error state.
' Unsupported operations raise VBAsh_STREAM_ERROR_NOT_SUPPORTED; operations on
' a closed stream raise VBAsh_STREAM_ERROR_CLOSED. Read, write, and open failures
' have separate reserved codes: VBAsh_STREAM_ERROR_READ_FAILURE,
' VBAsh_STREAM_ERROR_WRITE_FAILURE, and VBAsh_STREAM_ERROR_OPEN_FAILURE.
' CanRead reports only read capability; IsClosed reports only closed state;
' HasError reports only whether an operation failed.
' AtEndOfStream is meaningful only on open, readable streams. It returns False
' for non-readable or closed streams; otherwise it reports data exhaustion.
' ReadText returns at most maxChars; a non-positive limit returns empty text.
' WriteText returns the number of characters accepted. CloseStream is idempotent.
' This contract is text-based and counts characters; it is not a binary stream API.
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
