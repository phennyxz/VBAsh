Option Explicit

Public Function Parse(ByVal tokens As Collection, ByVal originalText As String) As clsCommand
    Dim command As clsCommand
    Dim token As clsToken
    Dim hasName As Boolean

    Set command = New clsCommand
    hasName = False

    For Each token In tokens
        Select Case token.TokenType
            Case "EOF"
                Exit For
            Case "WORD"
                If Not hasName Then
                    command.Init token.Value, originalText
                    hasName = True
                Else
                    command.AddArg token.Value
                End If
            Case Else
                Err.Raise vbObjectError + 1001, "clsParser.Parse", _
                    "Operador ainda não suportado: " & token.Value
        End Select
    Next token

    If Not hasName Then
        Err.Raise vbObjectError + 1002, "clsParser.Parse", "Comando vazio."
    End If

    Set Parse = command
End Function