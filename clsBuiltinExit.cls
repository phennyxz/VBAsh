Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim exitCode As Long
    Dim numericValue As Double

    If args.Count > 1 Then Err.Raise vbObjectError + 1022, "clsBuiltinExit", "Uso: exit [status]"
    If args.Count = 1 Then
        If Not IsNumeric(CStr(args(1))) Then Err.Raise vbObjectError + 1023, "clsBuiltinExit", "Status inválido."
        numericValue = CDbl(args(1))
        If numericValue < 0 Or numericValue > 255 Or numericValue <> Fix(numericValue) Then
            Err.Raise vbObjectError + 1024, "clsBuiltinExit", "Status deve ser um inteiro entre 0 e 255."
        End If
        exitCode = CLng(numericValue)
    End If

    session.LastExitCode = exitCode
    session.ShouldExit = True
End Function
