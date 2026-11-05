' @VBAshRevision clsNullStream-r1
Option Explicit
Implements IStream

' Identifica a revisão carregada; não altera o comportamento do stream.
Public Property Get ModuleRevision() As String
    ModuleRevision = "clsNullStream-r1"
End Property

Private mCanRead As Boolean
Private mCanWrite As Boolean
Private mIsClosed As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

Private Sub Class_Initialize()
    Initialize
End Sub

' Initialize restores capabilities, the open state, and a clear error state.
' A null stream never stores written text.
Public Sub Initialize(Optional ByVal canRead As Boolean = True, _
                      Optional ByVal canWrite As Boolean = True)
    mCanRead = canRead
    mCanWrite = canWrite
    mIsClosed = False
    ClearStreamError
End Sub

Private Property Get IStream_CanRead() As Boolean
    IStream_CanRead = mCanRead
End Property

Private Property Get IStream_CanWrite() As Boolean
    IStream_CanWrite = mCanWrite
End Property

Private Property Get IStream_AtEndOfStream() As Boolean
    IStream_AtEndOfStream = mCanRead And Not mIsClosed
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
    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanRead Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support reading."
    End If
    ClearStreamError
    IStream_ReadText = ""
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    If mIsClosed Then
        RaiseStreamError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    End If
    If Not mCanWrite Then
        RaiseStreamError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Stream does not support writing."
    End If
    ClearStreamError
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mIsClosed = True
    ClearStreamError
End Sub

Private Sub RaiseStreamError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsNullStream", description
End Sub

Private Sub ClearStreamError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
