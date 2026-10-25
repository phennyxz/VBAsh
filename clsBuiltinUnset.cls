Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim i As Long

    If args.Count = 0 Then Err.Raise vbObjectError + 1019, "clsBuiltinUnset", "Uso: unset NOME [NOME ...]"
    For i = 1 To args.Count
        If Not session.IsValidEnvironmentName(CStr(args(i))) Then
            Err.Raise vbObjectError + 1020, "clsBuiltinUnset", "Nome de variável inválido: " & CStr(args(i))
        End If
        session.RemoveEnv CStr(args(i))
    Next i
End Function
