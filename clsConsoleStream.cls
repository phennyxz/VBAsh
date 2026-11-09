' @VBAshRevision clsConsoleStream-r1
Option Explicit
Implements IStream

Private mBuffer As String
Private mPosition As Long
Private mCanRead As Boolean
Private mCanWrite As Boolean
Private mIsClosed As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsConsoleStream-r1"
End Property

Private Sub Class_Initialize()
    Initialize
End Sub

' Host-facing capture adapter. The UserForm consumes the returned text after
' ProcessInput; this class deliberately does not display modal UI itself.
Public Sub Initialize(Optional ByVal canRead As Boolean = False, _
                      Optional ByVal canWrite As Boolean = True)
    mBuffer = ""
    mPosition = 1
    mCanRead = canRead
    mCanWrite = canWrite
    mIsClosed = False
    ClearError
End Sub

Public Sub Reset()
    Initialize mCanRead, mCanWrite
End Sub

Public Property Get Contents() As String
    Contents = mBuffer
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
    Dim count As Long
    If mIsClosed Then RaiseError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    If Not mCanRead Then RaiseError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Console stream is not readable."
    ClearError
    If maxChars <= 0 Then Exit Function
    count = Len(mBuffer) - mPosition + 1
    If count <= 0 Then Exit Function
    If count > maxChars Then count = maxChars
    IStream_ReadText = Mid$(mBuffer, mPosition, count)
    mPosition = mPosition + count
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    If mIsClosed Then RaiseError VBAsh_STREAM_ERROR_CLOSED, "Stream is closed."
    If Not mCanWrite Then RaiseError VBAsh_STREAM_ERROR_NOT_SUPPORTED, "Console stream is not writable."
    ClearError
    mBuffer = mBuffer & text
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mIsClosed = True
    ClearError
End Sub

Private Sub RaiseError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsConsoleStream", description
End Sub

Private Sub ClearError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
