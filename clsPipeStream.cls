' @VBAshRevision clsPipeStream-r1
Option Explicit
Implements IStream

Private mBuffer As String
Private mReadPosition As Long
Private mWriterOpen As Boolean
Private mReaderOpen As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsPipeStream-r1"
End Property

Private Sub Class_Initialize()
    Initialize
End Sub

' Buffered text pipe primitive. Process scheduling, blocking reads, endpoint
' reference counting, and pipeline status remain the responsibility of Phase 9.
Public Sub Initialize()
    mBuffer = ""
    mReadPosition = 1
    mWriterOpen = True
    mReaderOpen = True
    ClearError
End Sub

Public Property Get BufferedLength() As Long
    BufferedLength = Len(mBuffer) - mReadPosition + 1
    If BufferedLength < 0 Then BufferedLength = 0
End Property

Private Property Get IStream_CanRead() As Boolean
    IStream_CanRead = mReaderOpen
End Property

Private Property Get IStream_CanWrite() As Boolean
    IStream_CanWrite = mWriterOpen
End Property

Private Property Get IStream_AtEndOfStream() As Boolean
    IStream_AtEndOfStream = mReaderOpen And Not mWriterOpen And mReadPosition > Len(mBuffer)
End Property

Private Property Get IStream_IsClosed() As Boolean
    IStream_IsClosed = Not mReaderOpen And Not mWriterOpen
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
    If Not mReaderOpen Then RaiseError VBAsh_STREAM_ERROR_CLOSED, "Pipe reader is closed."
    ClearError
    If maxChars <= 0 Then Exit Function
    count = Len(mBuffer) - mReadPosition + 1
    If count <= 0 Then Exit Function
    If count > maxChars Then count = maxChars
    IStream_ReadText = Mid$(mBuffer, mReadPosition, count)
    mReadPosition = mReadPosition + count
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    If Not mWriterOpen Then RaiseError VBAsh_STREAM_ERROR_CLOSED, "Pipe writer is closed."
    ClearError
    mBuffer = mBuffer & text
    IStream_WriteText = Len(text)
End Function

Private Sub IStream_CloseStream()
    mReaderOpen = False
    mWriterOpen = False
    ClearError
End Sub

Public Sub CloseWriter()
    mWriterOpen = False
End Sub

Public Sub CloseReader()
    mReaderOpen = False
End Sub

Private Sub RaiseError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsPipeStream", description
End Sub

Private Sub ClearError()
    mLastErrorNumber = 0
    mLastErrorDescription = ""
End Sub
