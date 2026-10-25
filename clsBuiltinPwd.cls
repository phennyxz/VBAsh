Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    If args.Count > 0 Then Err.Raise vbObjectError + 1011, "clsBuiltinPwd", "Uso: pwd"
    IBuiltin_Execute = session.CurrentDirectory
    session.HasOutput = True
End Function
