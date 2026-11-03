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
    mPosition = 1
    mCanRead = True
    mCanWrite = True
End Sub

Public Sub ConfigureCapabilities(ByVal canRead As Boolean, ByVal canWrite As Boolean)
    mCanRead = canRead
    mCanWrite = canWrite
End Sub

Private Property Get IStream_CanRead() As Boolean
    IStream_CanRead = mCanRead
End Property

Private Property Get IStream_CanWrite() As Boolean
    IStream_CanWrite = mCanWrite
End Property

Private Property Get IStream_AtEndOfStream() As Boolean
    IStream_AtEndOfStream = mIsClosed Or Not mCanRead Or mPosition > Len(mBuffer)
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

    If mIsClosed Then RaiseStreamError "Stream is closed."
    If Not mCanRead Then RaiseStreamError "Stream does not support reading."
    ClearStreamError
    If maxChars <= 0 Then Exit Function

    remainingChars = Len(mBuffer) - mPosition + 1
    If remainingChars <= 0 Then Exit Function
    If remainingChars > maxChars Then remainingChars = maxChars
    IStream_ReadText = Mid$(mBuffer, mPosition, remainingChars)
    mPosition = mPosition + remainingChars
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    If mIsClosed Then RaiseStreamError "Stream is closed."
    If Not mCanWrite Then RaiseStreamError "Stream does not support writing."
    ClearStreamError
    mBuffer = mBuffer & text
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mIsClosed = True
    ClearStreamError
End Sub

Private Sub RaiseStreamError(ByVal description As String)
    mLastErrorNumber = VBAsh_STREAM_ERROR_NOT_SUPPORTED
    mLastErrorDescription = description
    Err.Raise VBAsh_STREAM_ERROR_NOT_SUPPORTED, "clsStreamContractProbe", description
End Sub

Private Sub ClearStreamError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
