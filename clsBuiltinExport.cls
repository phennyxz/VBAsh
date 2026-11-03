Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim i As Long
    Dim key As Variant
    Dim assignment As String
    Dim separator As Long

    If args.Count = 0 Then
        For Each key In session.Env.Keys
            If Len(IBuiltin_Execute) > 0 Then IBuiltin_Execute = IBuiltin_Execute & vbCrLf
            IBuiltin_Execute = IBuiltin_Execute & CStr(key) & "=" & CStr(session.Env(key))
        Next key
        session.HasOutput = (Len(IBuiltin_Execute) > 0)
        Exit Function
    End If

    For i = 1 To args.Count
        assignment = CStr(args(i))
        separator = InStr(1, assignment, "=", vbBinaryCompare)
        If separator > 0 Then
            session.SetEnv Left$(assignment, separator - 1), Mid$(assignment, separator + 1)
        ElseIf Not session.Env.Exists(assignment) Then
            Err.Raise vbObjectError + 1018, "clsBuiltinExport", "Variável não definida: " & assignment
        End If
    Next i
End Function
