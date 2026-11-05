' @VBAshRevision clsFileDescriptorTable-r1
Option Explicit

Private Const FD_ERROR_NOT_INITIALIZED As Long = vbObjectError + 1250
Private Const FD_ERROR_INVALID As Long = vbObjectError + 1251
Private Const FD_ERROR_NOT_FOUND As Long = vbObjectError + 1252
Private Const FD_ERROR_NO_FREE_DESCRIPTOR As Long = vbObjectError + 1253
Private Const FD_ERROR_ALREADY_INITIALIZED As Long = vbObjectError + 1254

Private mDescriptors As Object
Private mOwnedStreams As Collection
Private mIsInitialized As Boolean

' Identifies the table API and ownership implementation revision.
Public Property Get ModuleRevision() As String
    ModuleRevision = "clsFileDescriptorTable-r1"
End Property

Private Sub Class_Initialize()
    Set mDescriptors = CreateObject("Scripting.Dictionary")
    Set mOwnedStreams = New Collection
End Sub

' Standard streams are borrowed unless ownership is explicitly transferred.
Public Sub Initialize(ByVal stdin As IStream, ByVal stdout As IStream, _
                      ByVal stderr As IStream, _
                      Optional ByVal ownsStandardStreams As Boolean = False)
    If mIsInitialized Then
        RaiseTableError FD_ERROR_ALREADY_INITIALIZED, "Initialize", _
                        "The descriptor table is already initialized."
    End If
    If (stdin Is Nothing) Or (stdout Is Nothing) Or (stderr Is Nothing) Then
        RaiseTableError FD_ERROR_INVALID, "Initialize", _
                        "Standard descriptors require non-Nothing IStream objects."
    End If

    Set mDescriptors.Item("0") = stdin
    Set mDescriptors.Item("1") = stdout
    Set mDescriptors.Item("2") = stderr
    mIsInitialized = True

    If ownsStandardStreams Then
        AddOwnedStream stdin
        AddOwnedStream stdout
        AddOwnedStream stderr
    End If
End Sub

Public Function HasDescriptor(ByVal descriptor As Long) As Boolean
    EnsureInitialized "HasDescriptor"
    ValidateDescriptor descriptor, "HasDescriptor"
    HasDescriptor = mDescriptors.Exists(CStr(descriptor))
End Function

Public Function GetDescriptor(ByVal descriptor As Long) As IStream
    EnsureInitialized "GetDescriptor"
    ValidateDescriptor descriptor, "GetDescriptor"
    If Not mDescriptors.Exists(CStr(descriptor)) Then
        RaiseTableError FD_ERROR_NOT_FOUND, "GetDescriptor", _
                        "Descriptor " & CStr(descriptor) & " does not exist."
    End If
    Set GetDescriptor = mDescriptors.Item(CStr(descriptor))
End Function

' Ownership is acquired by stream identity and cannot be revoked by a later
' borrowed association of the same object.
Public Sub SetDescriptor(ByVal descriptor As Long, ByVal stream As IStream, _
                         Optional ByVal ownsStream As Boolean = False)
    EnsureInitialized "SetDescriptor"
    ValidateDescriptor descriptor, "SetDescriptor"
    If stream Is Nothing Then
        RaiseTableError FD_ERROR_INVALID, "SetDescriptor", _
                        "A descriptor cannot be associated with Nothing."
    End If
    ReplaceDescriptor descriptor, stream, ownsStream, "SetDescriptor"
End Sub

' Duplicated descriptors reference the same IStream object and share its
' object-level ownership state.
Public Sub DuplicateDescriptor(ByVal sourceDescriptor As Long, _
                               ByVal destinationDescriptor As Long)
    Dim stream As IStream

    EnsureInitialized "DuplicateDescriptor"
    ValidateDescriptor sourceDescriptor, "DuplicateDescriptor"
    ValidateDescriptor destinationDescriptor, "DuplicateDescriptor"
    If Not mDescriptors.Exists(CStr(sourceDescriptor)) Then
        RaiseTableError FD_ERROR_NOT_FOUND, "DuplicateDescriptor", _
                        "Source descriptor " & CStr(sourceDescriptor) & " does not exist."
    End If
    If sourceDescriptor = destinationDescriptor Then Exit Sub

    Set stream = mDescriptors.Item(CStr(sourceDescriptor))
    ReplaceDescriptor destinationDescriptor, stream, False, "DuplicateDescriptor"
End Sub

' A failed final close leaves the descriptor and ownership association intact,
' so the caller can inspect the stream and retry.
Public Sub CloseDescriptor(ByVal descriptor As Long)
    Dim stream As IStream
    Dim key As String

    EnsureInitialized "CloseDescriptor"
    ValidateDescriptor descriptor, "CloseDescriptor"
    key = CStr(descriptor)
    If Not mDescriptors.Exists(key) Then
        RaiseTableError FD_ERROR_NOT_FOUND, "CloseDescriptor", _
                        "Descriptor " & key & " does not exist."
    End If

    Set stream = mDescriptors.Item(key)
    If Not HasOtherDescriptorReference(stream, key) Then
        If IsOwnedStream(stream) Then
            CloseOwnedStream stream, descriptor, "CloseDescriptor"
            RemoveOwnedStream stream
        End If
    End If
    mDescriptors.Remove key
End Sub

' Returns the lowest unused descriptor at or above 3 without reserving it.
Public Function FindFreeDescriptor() As Long
    Dim descriptor As Long

    EnsureInitialized "FindFreeDescriptor"
    descriptor = 3
    Do While mDescriptors.Exists(CStr(descriptor))
        If descriptor = 2147483647 Then
            RaiseTableError FD_ERROR_NO_FREE_DESCRIPTOR, "FindFreeDescriptor", _
                            "No free descriptor remains in the supported Long range."
        End If
        descriptor = descriptor + 1
    Loop
    FindFreeDescriptor = descriptor
End Function

Private Sub ReplaceDescriptor(ByVal descriptor As Long, ByVal stream As IStream, _
                              ByVal ownsStream As Boolean, ByVal operation As String)
    Dim key As String
    Dim previousStream As IStream

    key = CStr(descriptor)
    If mDescriptors.Exists(key) Then
        Set previousStream = mDescriptors.Item(key)
        If previousStream Is stream Then
            If ownsStream Then AddOwnedStream stream
            Exit Sub
        End If

        If Not HasOtherDescriptorReference(previousStream, key) Then
            If IsOwnedStream(previousStream) Then
                CloseOwnedStream previousStream, descriptor, operation
                RemoveOwnedStream previousStream
            End If
        End If
    End If

    Set mDescriptors.Item(key) = stream
    If ownsStream Then AddOwnedStream stream
End Sub

Private Function HasOtherDescriptorReference(ByVal stream As IStream, _
                                             ByVal excludedKey As String) As Boolean
    Dim descriptorKey As Variant
    Dim candidate As IStream

    For Each descriptorKey In mDescriptors.Keys
        If CStr(descriptorKey) <> excludedKey Then
            Set candidate = mDescriptors.Item(CStr(descriptorKey))
            If candidate Is stream Then
                HasOtherDescriptorReference = True
                Exit Function
            End If
        End If
    Next descriptorKey
End Function

Private Function IsOwnedStream(ByVal stream As IStream) As Boolean
    Dim index As Long
    Dim ownedStream As IStream

    For index = 1 To mOwnedStreams.Count
        Set ownedStream = mOwnedStreams(index)
        If ownedStream Is stream Then
            IsOwnedStream = True
            Exit Function
        End If
    Next index
End Function

Private Sub AddOwnedStream(ByVal stream As IStream)
    If Not IsOwnedStream(stream) Then mOwnedStreams.Add stream
End Sub

Private Sub RemoveOwnedStream(ByVal stream As IStream)
    Dim index As Long
    Dim ownedStream As IStream

    For index = mOwnedStreams.Count To 1 Step -1
        Set ownedStream = mOwnedStreams(index)
        If ownedStream Is stream Then
            mOwnedStreams.Remove index
            Exit Sub
        End If
    Next index
End Sub

Private Sub CloseOwnedStream(ByVal stream As IStream, ByVal descriptor As Long, _
                             ByVal operation As String)
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String
    Dim originalErrorSource As String

    On Error GoTo CloseFailure
    stream.CloseStream
    Exit Sub

CloseFailure:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    originalErrorSource = Err.Source
    On Error GoTo 0
    If originalErrorNumber = 0 Then originalErrorNumber = FD_ERROR_INVALID
    Err.Raise originalErrorNumber, "clsFileDescriptorTable." & operation, _
              "Could not close the owned stream for descriptor " & CStr(descriptor) & _
              " (original source [" & originalErrorSource & "]): " & _
              originalErrorDescription
End Sub

Private Sub EnsureInitialized(ByVal operation As String)
    If Not mIsInitialized Then
        RaiseTableError FD_ERROR_NOT_INITIALIZED, operation, _
                        "Initialize the standard descriptors before using the table."
    End If
End Sub

Private Sub ValidateDescriptor(ByVal descriptor As Long, ByVal operation As String)
    If descriptor < 0 Then
        RaiseTableError FD_ERROR_INVALID, operation, _
                        "Descriptor numbers must be non-negative."
    End If
End Sub

Private Sub RaiseTableError(ByVal errorNumber As Long, ByVal operation As String, _
                            ByVal description As String)
    Err.Raise errorNumber, "clsFileDescriptorTable." & operation, description
End Sub
