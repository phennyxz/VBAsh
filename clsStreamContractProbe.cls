Option Explicit
Implements IStream

Private mBuffer As String
Private mPosition As Long
Private mCanRead As Boolean
Private mCanWrite As Boolean
Private mIsClosed As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String
Private mFailClose As Boolean
Private mCloseCount As Long
Private mCloseBeforeFailure As Boolean
Private mCloseErrorNumber As Long
Private mCloseErrorSource As String
Private mCloseErrorDescription As String
Private mCleanupReentryTable As clsFileDescriptorTable

Private Sub Class_Initialize()
    mPosition = 1
    mCanRead = True
    mCanWrite = True
    mCloseErrorNumber = vbObjectError + 1260
    mCloseErrorSource = "clsStreamContractProbe.CloseStream"
    mCloseErrorDescription = "Configured close failure."
End Sub

Public Sub ConfigureCapabilities(ByVal canRead As Boolean, ByVal canWrite As Boolean)
    mCanRead = canRead
    mCanWrite = canWrite
End Sub

Public Sub ConfigureCloseFailure(ByVal shouldFail As Boolean)
    mFailClose = shouldFail
End Sub

Public Sub ConfigureCloseFailureDetails(ByVal errorNumber As Long, ByVal errorSource As String, _
                                        ByVal errorDescription As String, _
                                        ByVal closeBeforeFailure As Boolean)
    mCloseErrorNumber = errorNumber
    mCloseErrorSource = errorSource
    mCloseErrorDescription = errorDescription
    mCloseBeforeFailure = closeBeforeFailure
    mFailClose = True
End Sub

Public Sub ConfigureCleanupReentry(ByVal table As clsFileDescriptorTable)
    Set mCleanupReentryTable = table
End Sub

Public Property Get CloseCount() As Long
    CloseCount = mCloseCount
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
    mBuffer = mBuffer & text
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mCloseCount = mCloseCount + 1
    If Not mCleanupReentryTable Is Nothing Then mCleanupReentryTable.CloseAll
    If mFailClose Then
        If mCloseBeforeFailure Then mIsClosed = True
        Err.Raise mCloseErrorNumber, mCloseErrorSource, mCloseErrorDescription
    End If
    mIsClosed = True
    ClearStreamError
End Sub

Private Sub RaiseStreamError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsStreamContractProbe", description
End Sub

Private Sub ClearStreamError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
