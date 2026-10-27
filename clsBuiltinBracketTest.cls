Option Explicit
Implements IBuiltin

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim testBuiltin As IBuiltin
    Dim testArgs As New Collection
    Dim i As Long

    If args.Count = 0 Then Err.Raise vbObjectError + 1062, "clsBuiltinBracketTest", "Uso: [ expressão ]"
    If StrComp(CStr(args(args.Count)), "]", vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 1063, "clsBuiltinBracketTest", "Esperado ']' ao final do teste."
    End If

    For i = 1 To args.Count - 1
        testArgs.Add args(i)
    Next i

    Set testBuiltin = New clsBuiltinTest
    IBuiltin_Execute = testBuiltin.Execute(session, testArgs)
End Function
