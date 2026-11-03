Option Explicit
Implements IStream

Private mBuffer As String
Private mPosition As Long
Private mCanRead As Boolean
Private mCanWrite As Boolean
Private mIsClosed As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

Private Sub Class_Initialize()
    Initialize
End Sub

' Initialize replaces the buffer and resets the cursor, capabilities, and error state.
Public Sub Initialize(Optional ByVal initialContent As String = "", _
                      Optional ByVal canRead As Boolean = True, _
                      Optional ByVal canWrite As Boolean = True)
    mBuffer = initialContent
    mPosition = 1
    mCanRead = canRead
    mCanWrite = canWrite
    mIsClosed = False
    ClearStreamError
End Sub

Public Property Get Contents() As String
    Contents = mBuffer
End Property

' Position is the 1-based position of the next character to read.
Public Property Get Position() As Long
    Position = mPosition
End Property

Private Property Get IStream_CanRead() As Boolean
    IStream_CanRead = mCanRead
End Property

Private Property Get IStream_CanWrite() As Boolean
    IStream_CanWrite = mCanWrite
End Property

Private Property Get IStream_AtEndOfStream() As Boolean
    IStream_AtEndOfStream = mCanRead And Not mIsClosed And mPosition > Len(mBuffer)
End Property

Private Property Get IStream_IsClosed() As Boolean
    IStream_IsClosed = mIsClosed
End Property

Private Property Get IStream_HasError() As Boolean
    IStream_HasError = (mLastErrorNumber <> 0)
End Property

Private Property Get IStream_LastErrorNumber() As Long
    IStream_LastErrorNumber = mLastErrorNumber
End Property

Private Property Get IStream_LastErrorDescription() As String
    IStream_LastErrorDescription = mLastErrorDescription
End Property

Private Function IStream_ReadText(ByVal maxChars As Long) As String
    Dim remainingChars As Long

    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanRead Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support reading."
    End If
    ClearStreamError
    If maxChars <= 0 Then Exit Function

    remainingChars = Len(mBuffer) - mPosition + 1
    If remainingChars <= 0 Then Exit Function
    If remainingChars > maxChars Then remainingChars = maxChars
    IStream_ReadText = Mid$(mBuffer, mPosition, remainingChars)
    mPosition = mPosition + remainingChars
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanWrite Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support writing."
    End If
    ClearStreamError
    ' Writes append to the buffer without moving the read position.
    mBuffer = mBuffer & text
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mIsClosed = True
    ClearStreamError
End Sub

Private Sub RaiseStreamError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsMemoryStream", description
End Sub

Private Sub ClearStreamError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
