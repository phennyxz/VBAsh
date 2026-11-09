' @VBAshRevision clsBuiltinAlias-r1
Option Explicit
Implements IBuiltin

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsBuiltinAlias-r1"
End Property

Private Function IBuiltin_Execute(ByVal session As clsSession, ByVal args As Collection) As String
    Dim aliasName As String
    Dim aliasValue As String
    Dim separator As Long
    Dim names As Collection
    Dim item As Variant
    Dim output As String
    Dim i As Long

    If args.Count = 0 Then
        Set names = session.GetAliasNames()
        For Each item In names
            AppendAlias output, CStr(item), session.GetAlias(CStr(item))
        Next item
    Else
        For i = 1 To args.Count
            aliasValue = CStr(args(i))
            separator = InStr(1, aliasValue, "=", vbBinaryCompare)
            If separator > 0 Then
                aliasName = Left$(aliasValue, separator - 1)
                session.RegisterAlias aliasName, Mid$(aliasValue, separator + 1), True
            ElseIf session.HasAlias(aliasValue) Then
                AppendAlias output, aliasValue, session.GetAlias(aliasValue)
            End If
        Next i
    End If

    IBuiltin_Execute = output
    session.HasOutput = (Len(output) > 0)
End Function

Private Sub AppendAlias(ByRef output As String, ByVal aliasName As String, ByVal aliasValue As String)
    If Len(output) > 0 Then output = output & vbCrLf
    output = output & "alias " & aliasName & "=" & ShellQuote(aliasValue)
End Sub

Private Function ShellQuote(ByVal value As String) As String
    Dim escapedQuote As String
    escapedQuote = Chr$(39) & "\" & Chr$(39) & Chr$(39)
    ShellQuote = Chr$(39) & Replace(value, Chr$(39), escapedQuote) & Chr$(39)
End Function
