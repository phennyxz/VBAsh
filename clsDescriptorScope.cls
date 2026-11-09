' @VBAshRevision clsDescriptorScope-r1
Option Explicit

Private mEntries As Collection
Private mIsRestored As Boolean
Private mIsRestoring As Boolean

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsDescriptorScope-r1"
End Property

Private Sub Class_Initialize()
    Set mEntries = New Collection
End Sub

Public Property Get IsRestored() As Boolean
    IsRestored = mIsRestored
End Property

Public Property Get IsRestoring() As Boolean
    IsRestoring = mIsRestoring
End Property

Public Property Get EntryCount() As Long
    EntryCount = mEntries.Count
End Property

Public Function GetEntry(ByVal index As Long) As clsDescriptorScopeEntry
    If index < 1 Or index > mEntries.Count Then
        Err.Raise vbObjectError + 1260, "clsDescriptorScope.GetEntry", _
                  "Índice de snapshot fora do intervalo."
    End If
    Set GetEntry = mEntries(index)
End Function

Public Sub AddEntry(ByVal entry As clsDescriptorScopeEntry)
    If mIsRestored Or mIsRestoring Then
        Err.Raise vbObjectError + 1261, "clsDescriptorScope.AddEntry", _
                  "O snapshot não aceita novas entradas neste estado."
    End If
    If entry Is Nothing Then
        Err.Raise vbObjectError + 1262, "clsDescriptorScope.AddEntry", _
                  "Entrada de snapshot ausente."
    End If
    mEntries.Add entry
End Sub

Public Sub BeginRestore()
    If mIsRestored Then Exit Sub
    If mIsRestoring Then
        Err.Raise vbObjectError + 1263, "clsDescriptorScope.BeginRestore", _
                  "A restauração já está em andamento."
    End If
    mIsRestoring = True
End Sub

Public Sub CompleteRestore()
    mIsRestoring = False
    mIsRestored = True
End Sub

Public Sub CancelRestoreAttempt()
    mIsRestoring = False
End Sub
