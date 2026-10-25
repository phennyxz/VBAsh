Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    If args.Count > 0 Then Err.Raise vbObjectError + 1021, "clsBuiltinHelp", "Uso: help"
    IBuiltin_Execute = "Builtins: echo [-n], pwd, cd [diretório], export, setenv, unset, help, exit [status]"
    session.HasOutput = True
End Function
