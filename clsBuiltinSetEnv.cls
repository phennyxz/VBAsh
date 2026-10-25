Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    If args.Count <> 2 Then Err.Raise vbObjectError + 1017, "clsBuiltinSetEnv", "Uso: setenv NOME valor"
    session.SetEnv CStr(args(1)), CStr(args(2))
End Function
