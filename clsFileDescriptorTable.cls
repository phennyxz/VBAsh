' @VBAshRevision clsFileDescriptorTable-r2
Option Explicit

Private Const FD_ERROR_NOT_INITIALIZED As Long = vbObjectError + 1250
Private Const FD_ERROR_INVALID As Long = vbObjectError + 1251
Private Const FD_ERROR_NOT_FOUND As Long = vbObjectError + 1252
Private Const FD_ERROR_NO_FREE_DESCRIPTOR As Long = vbObjectError + 1253
Private Const FD_ERROR_ALREADY_INITIALIZED As Long = vbObjectError + 1254
Private Const FD_ERROR_CLOSE_ALL As Long = vbObjectError + 1255
Private Const FD_ERROR_CLEANUP_IN_PROGRESS As Long = vbObjectError + 1256

Private mDescriptors As Object
Private mOwnedStreams As Collection
Private mIsInitialized As Boolean
Private mClosingAll As Boolean
Private mScopes As Collection

' Identifies the table API and ownership implementation revision.
Public Property Get ModuleRevision() As String
    ModuleRevision = "clsFileDescriptorTable-r2"
End Property

Private Sub Class_Initialize()
    Set mDescriptors = CreateObject("Scripting.Dictionary")
    Set mOwnedStreams = New Collection
    Set mScopes = New Collection
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
    EnsureNotClosingAll "SetDescriptor"
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
    EnsureNotClosingAll "DuplicateDescriptor"
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
    EnsureNotClosingAll "CloseDescriptor"
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

' The caller must explicitly clean up before releasing the table. Ownership
' and alias counts are local to this instance, not to other tables or callers.
Public Sub CloseAll()
    Dim keys As Variant
    Dim streams As Collection
    Dim stream As IStream
    Dim candidate As IStream
    Dim processed() As Boolean
    Dim index As Long
    Dim other As Long
    Dim count As Long
    Dim descriptorList As String
    Dim failures As String
    Dim closeNumber As Long
    Dim closeSource As String
    Dim closeDescription As String
    Dim unexpectedNumber As Long
    Dim unexpectedSource As String
    Dim unexpectedDescription As String

    EnsureInitialized "CloseAll"
    EnsureNotClosingAll "CloseAll"
    EnsureNoActiveScopes "CloseAll"
    count = mDescriptors.Count
    If count = 0 Then Exit Sub

    keys = mDescriptors.Keys
    Set streams = New Collection
    ReDim processed(0 To count - 1)
    For index = 0 To count - 1
        Set stream = mDescriptors.Item(CStr(keys(index)))
        streams.Add stream
    Next index

    mClosingAll = True
    On Error GoTo UnexpectedFailure
    For index = 0 To count - 1
        If Not processed(index) Then
            Set stream = streams.Item(index + 1)
            descriptorList = ""
            For other = index To count - 1
                Set candidate = streams.Item(other + 1)
                If candidate Is stream Then
                    processed(other) = True
                    If Len(descriptorList) > 0 Then descriptorList = descriptorList & ", "
                    descriptorList = descriptorList & CStr(keys(other))
                End If
            Next other

            closeNumber = 0
            closeSource = ""
            closeDescription = ""
            If IsOwnedStream(stream) Then
                TryCloseForCleanup stream, closeNumber, closeSource, closeDescription
            End If
            If closeNumber <> 0 Then
                If Len(failures) > 0 Then failures = failures & vbCrLf
                failures = failures & "FDs [" & descriptorList & "]: error " & _
                           CStr(closeNumber) & "; source [" & closeSource & _
                           "]; description [" & closeDescription & "]"
            Else
                For other = index To count - 1
                    Set candidate = streams.Item(other + 1)
                    If candidate Is stream Then mDescriptors.Remove CStr(keys(other))
                Next other
                RemoveOwnedStream stream
            End If
        End If
    Next index
    mClosingAll = False
    On Error GoTo 0
    If Len(failures) > 0 Then
        Err.Raise FD_ERROR_CLOSE_ALL, "clsFileDescriptorTable.CloseAll", failures
    End If
    Exit Sub

UnexpectedFailure:
    unexpectedNumber = Err.Number
    unexpectedSource = Err.Source
    unexpectedDescription = Err.Description
    mClosingAll = False
    On Error GoTo 0
    If Len(failures) > 0 Then
        unexpectedDescription = unexpectedDescription & vbCrLf & _
                                "Earlier close failures:" & vbCrLf & failures
    End If
    Err.Raise unexpectedNumber, unexpectedSource, unexpectedDescription
End Sub

' Only stream close errors are collected; other errors propagate normally.
Private Sub TryCloseForCleanup(ByVal stream As IStream, ByRef errorNumber As Long, _
                               ByRef errorSource As String, ByRef errorDescription As String)
    On Error GoTo CloseFailure
    stream.CloseStream
    Exit Sub
CloseFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
End Sub


' Captures the current descriptor bindings and pins their stream identities.
Public Function BeginDescriptorScope() As clsDescriptorScope
    Dim scope As clsDescriptorScope
    Dim entry As clsDescriptorScopeEntry
    Dim keys As Variant
    Dim index As Long
    Dim stream As IStream

    EnsureInitialized "BeginDescriptorScope"
    EnsureNotClosingAll "BeginDescriptorScope"
    Set scope = New clsDescriptorScope
    If mDescriptors.Count > 0 Then
        keys = mDescriptors.Keys
        For index = LBound(keys) To UBound(keys)
            Set stream = mDescriptors.Item(CStr(keys(index)))
            Set entry = New clsDescriptorScopeEntry
            entry.Descriptor = CLng(keys(index))
            entry.WasPresent = True
            Set entry.SavedStream = stream
            entry.WasOwned = IsOwnedStream(stream)
            scope.AddEntry entry
        Next index
    End If
    mScopes.Add scope
    Set BeginDescriptorScope = scope
End Function

' Restores a scope in strict LIFO order. Failed operations leave the scope
' active so the caller can retry without losing its saved stream references.
Public Sub RestoreDescriptorScope(ByVal scope As clsDescriptorScope)
    Dim entry As clsDescriptorScopeEntry
    Dim index As Long
    Dim keys As Variant
    Dim key As String
    Dim keep As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    EnsureInitialized "RestoreDescriptorScope"
    EnsureNotClosingAll "RestoreDescriptorScope"
    If scope Is Nothing Then
        RaiseTableError FD_ERROR_INVALID, "RestoreDescriptorScope", "Scope ausente."
    End If
    If mScopes.Count = 0 Then
        RaiseTableError FD_ERROR_INVALID, "RestoreDescriptorScope", "Não há escopos ativos."
    End If
    If Not (mScopes.Item(mScopes.Count) Is scope) Then
        RaiseTableError FD_ERROR_INVALID, "RestoreDescriptorScope", _
                        "Os escopos devem ser restaurados em ordem LIFO."
    End If
    If scope.IsRestored Then
        mScopes.Remove mScopes.Count
        Exit Sub
    End If

    scope.BeginRestore
    On Error GoTo RestoreFailure

    ' Remove descriptors created inside the scope.
    If mDescriptors.Count > 0 Then
        keys = mDescriptors.Keys
        For index = LBound(keys) To UBound(keys)
            key = CStr(keys(index))
            keep = ScopeContainsDescriptor(scope, CLng(key))
            If Not keep Then
                If mDescriptors.Exists(key) Then CloseDescriptor CLng(key)
            End If
        Next index
    End If

    ' Restore every binding that existed at scope entry.
    For index = 1 To scope.EntryCount
        Set entry = scope.GetEntry(index)
        If Not entry.RestoreCompleted Then
            If mDescriptors.Exists(CStr(entry.Descriptor)) Then
                SetDescriptor entry.Descriptor, entry.SavedStream, False
            Else
                SetDescriptor entry.Descriptor, entry.SavedStream, False
            End If
            entry.RestoreCompleted = True
        End If
    Next index

    scope.CompleteRestore
    mScopes.Remove mScopes.Count
    Exit Sub

RestoreFailure:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    scope.CancelRestoreAttempt
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Function ScopeContainsDescriptor(ByVal scope As clsDescriptorScope, _
                                         ByVal descriptor As Long) As Boolean
    Dim index As Long
    Dim entry As clsDescriptorScopeEntry
    For index = 1 To scope.EntryCount
        Set entry = scope.GetEntry(index)
        If entry.Descriptor = descriptor Then
            ScopeContainsDescriptor = True
            Exit Function
        End If
    Next index
End Function

Private Sub EnsureNoActiveScopes(ByVal operation As String)
    If mScopes.Count > 0 Then
        RaiseTableError FD_ERROR_CLEANUP_IN_PROGRESS, operation, _
                        "Não é possível limpar a tabela enquanto há escopos ativos."
    End If
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
    If IsPinnedStream(stream) Then HasOtherDescriptorReference = True
End Function

Private Function IsPinnedStream(ByVal stream As IStream) As Boolean
    Dim scopeIndex As Long
    Dim entryIndex As Long
    Dim scope As clsDescriptorScope
    Dim entry As clsDescriptorScopeEntry
    For scopeIndex = 1 To mScopes.Count
        Set scope = mScopes(scopeIndex)
        For entryIndex = 1 To scope.EntryCount
            Set entry = scope.GetEntry(entryIndex)
            If entry.WasPresent Then
                If entry.SavedStream Is stream Then
                    IsPinnedStream = True
                    Exit Function
                End If
            End If
        Next entryIndex
    Next scopeIndex
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

Private Sub EnsureNotClosingAll(ByVal operation As String)
    If mClosingAll Then
        RaiseTableError FD_ERROR_CLEANUP_IN_PROGRESS, operation, _
                        "The descriptor table is being cleaned up."
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
