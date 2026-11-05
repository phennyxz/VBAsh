Option Explicit
Implements IStream

Private mTextStream As Object
Private mCanRead As Boolean
Private mCanWrite As Boolean
Private mIsClosed As Boolean
Private mLastErrorNumber As Long
Private mLastErrorDescription As String

Private Sub Class_Initialize()
    mIsClosed = True
End Sub

Public Sub OpenForRead(ByVal path As String)
    Dim fileSystem As Object
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    PrepareForOpen
    On Error GoTo OpenFailure
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set mTextStream = fileSystem.OpenTextFile(path, 1, False)
    mCanRead = True
    mCanWrite = False
    mIsClosed = False
    ClearStreamError
    Exit Sub

OpenFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    Set mTextStream = Nothing
    mCanRead = False
    mCanWrite = False
    mIsClosed = True
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_OPEN_FAILURE, _
                     StreamFailureDescription("Open for read", originalErrorNumber, _
                                              originalErrorDescription)
End Sub

Public Sub OpenForWrite(ByVal path As String, Optional ByVal append As Boolean = False)
    Dim fileSystem As Object
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    PrepareForOpen
    On Error GoTo OpenFailure
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If append Then
        Set mTextStream = fileSystem.OpenTextFile(path, 8, True)
    Else
        Set mTextStream = fileSystem.CreateTextFile(path, True)
    End If
    mCanRead = False
    mCanWrite = True
    mIsClosed = False
    ClearStreamError
    Exit Sub

OpenFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    Set mTextStream = Nothing
    mCanRead = False
    mCanWrite = False
    mIsClosed = True
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_OPEN_FAILURE, _
                     StreamFailureDescription("Open for write", originalErrorNumber, _
                                              originalErrorDescription)
End Sub

Public Function ReadText(ByVal maxChars As Long) As String
    ReadText = ReadStreamText(maxChars)
End Function

Public Function WriteText(ByVal text As String) As Long
    WriteText = WriteStreamText(text)
End Function

Public Sub CloseStream()
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String
    Dim closeErrorNumber As Long

    If mTextStream Is Nothing Then
        mIsClosed = True
        ClearStreamError
        Exit Sub
    End If

    If mCanWrite Then
        closeErrorNumber = VBAsh_STREAM_ERROR_WRITE_FAILURE
    Else
        closeErrorNumber = VBAsh_STREAM_ERROR_READ_FAILURE
    End If
    On Error GoTo CloseFailure
    mTextStream.Close
    Set mTextStream = Nothing
    mIsClosed = True
    ClearStreamError
    Exit Sub

CloseFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    Set mTextStream = Nothing
    mIsClosed = True
    On Error GoTo 0
    RaiseStreamError closeErrorNumber, _
                     StreamFailureDescription("Close", originalErrorNumber, _
                                              originalErrorDescription)
End Sub

Private Property Get IStream_CanRead() As Boolean
    IStream_CanRead = mCanRead
End Property

Private Property Get IStream_CanWrite() As Boolean
    IStream_CanWrite = mCanWrite
End Property

Private Property Get IStream_AtEndOfStream() As Boolean
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    If mIsClosed Or Not mCanRead Then Exit Property

    On Error GoTo EndOfStreamFailure
    IStream_AtEndOfStream = mTextStream.AtEndOfStream
    Exit Property

EndOfStreamFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_READ_FAILURE, _
                     StreamFailureDescription("Check EOF", originalErrorNumber, _
                                              originalErrorDescription)
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
    IStream_ReadText = ReadStreamText(maxChars)
End Function

Private Function IStream_WriteText(ByVal text As String) As Long
    IStream_WriteText = WriteStreamText(text)
End Function

Private Sub IStream_CloseStream()
    CloseStream
End Sub

Private Function ReadStreamText(ByVal maxChars As Long) As String
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
    If mTextStream.AtEndOfStream Then Exit Function
    ReadStreamText = mTextStream.Read(maxChars)
    Exit Function

ReadFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_READ_FAILURE, _
                     StreamFailureDescription("Read", originalErrorNumber, _
                                              originalErrorDescription)
End Function

Private Function WriteStreamText(ByVal text As String) As Long
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
    mTextStream.Write text
    WriteStreamText = Len(text)
    Exit Function

WriteFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error GoTo 0
    RaiseStreamError VBAsh_STREAM_ERROR_WRITE_FAILURE, _
                     StreamFailureDescription("Write", originalErrorNumber, _
                                              originalErrorDescription)
End Function

Private Sub PrepareForOpen()
    If Not mTextStream Is Nothing Then CloseStream
    mIsClosed = True
    mCanRead = False
    mCanWrite = False
    ClearStreamError
End Sub

Private Sub RaiseStreamError(ByVal errorNumber As Long, ByVal description As String)
    mLastErrorNumber = errorNumber
    mLastErrorDescription = description
    Err.Raise errorNumber, "clsFileStream", description
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

Private Sub Class_Terminate()
    If Not mTextStream Is Nothing Then
        On Error Resume Next
        mTextStream.Close
        Set mTextStream = Nothing
        On Error GoTo 0
    End If
End Sub
