' @VBAshRevision clsBuiltinUnalias-r1
Option Explicit
Implements IBuiltin

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsBuiltinUnalias-r1"
End Property

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim i As Long
    Dim aliasName As String

    If args.Count = 0 Then
        Err.Raise vbObjectError + 1420, "clsBuiltinUnalias", "Uso: unalias [-a] NOME [NOME ...]"
    End If

    If CStr(args(1)) = "-a" Then
        If args.Count <> 1 Then
            Err.Raise vbObjectError + 1421, "clsBuiltinUnalias", "A opção -a não aceita outros argumentos."
        End If
        session.ClearAliases
        Exit Function
    End If

    For i = 1 To args.Count
        aliasName = CStr(args(i))
        If Left$(aliasName, 1) = "-" Then
            Err.Raise vbObjectError + 1422, "clsBuiltinUnalias", "Opção não suportada: " & aliasName
        End If
        If Not session.RemoveAlias(aliasName) Then
            Err.Raise vbObjectError + 1423, "clsBuiltinUnalias", "Alias não definido: " & aliasName
        End If
    Next i
End Function
