Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim firstArg As Long
    Dim i As Long

    firstArg = 1
    If args.Count > 0 Then
        If CStr(args(1)) = "-n" Then
            firstArg = 2
            session.SuppressNewline = True
        End If
    End If

    For i = firstArg To args.Count
        If i > firstArg Then IBuiltin_Execute = IBuiltin_Execute & " "
        IBuiltin_Execute = IBuiltin_Execute & CStr(args(i))
    Next i
    session.HasOutput = True
End Function
