' @VBAshRevision clsMemoryStream-r1
Option Explicit
Implements IStream

' Identifica a revisão carregada; não altera o comportamento do stream.
Public Property Get ModuleRevision() As String
    ModuleRevision = "clsMemoryStream-r1"
End Property

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
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanRead Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support reading."
    End If
    ClearStreamError
    If maxChars <= 0 Then Exit Function

    On Error GoTo ReadFailure
    remainingChars = Len(mBuffer) - mPosition + 1
    If remainingChars <= 0 Then Exit Function
    If remainingChars > maxChars Then remainingChars = maxChars
    IStream_ReadText = Mid$(mBuffer, mPosition, remainingChars)
    mPosition = mPosition + remainingChars
    Exit Function

ReadFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_READ_FAILURE, _
                     StreamFailureDescription("Read", originalErrorNumber, originalErrorDescription)
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    Dim newBuffer As String
    Dim writtenChars As Long
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanWrite Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support writing."
    End If
    ClearStreamError
    On Error GoTo WriteFailure
    writtenChars = Len(text)
    newBuffer = mBuffer & text
    ' Writes append to the buffer without moving the read position.
    IStream_WriteText = writtenChars
    mBuffer = newBuffer
    Exit Function

WriteFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_WRITE_FAILURE, _
                     StreamFailureDescription("Write", originalErrorNumber, originalErrorDescription)
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

Private Function StreamFailureDescription(ByVal operation As String, _
                                          ByVal originalErrorNumber As Long, _
                                          ByVal originalDescription As String) As String
    StreamFailureDescription = operation & " failed (error " & CStr(originalErrorNumber) & ")"
    If Len(originalDescription) > 0 Then
        StreamFailureDescription = StreamFailureDescription & ": " & originalDescription
    End If
End Function
