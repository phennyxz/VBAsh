Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    If args.Count > 0 Then Err.Raise vbObjectError + 1061, "clsBuiltinClear", "Uso: clear"
    session.ClearRequested = True
    session.HasOutput = False
End Function
