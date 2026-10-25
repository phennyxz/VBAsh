Option Explicit

Public Function Parse(ByVal tokens As Collection, ByVal originalText As String) As clsCommand
    Dim command As clsCommand
    Dim token As clsToken
    Dim redirection As clsRedirection
    Dim hasName As Boolean
    Dim expectsTarget As Boolean

    Set command = New clsCommand
    hasName = False

    For Each token In tokens
        Select Case token.TokenType
            Case "EOF"
                If expectsTarget Then
                    Err.Raise vbObjectError + 1003, "clsParser.Parse", "Destino de redirecionamento ausente."
                End If
                Exit For
            Case "WORD"
                If expectsTarget Then
                    redirection.Target = token.Value
                    command.AddRedirection redirection
                    expectsTarget = False
                ElseIf Not hasName Then
                    command.Init token.Value, originalText
                    hasName = True
                Else
                    command.AddArg token.Value
                End If
            Case "REDIR_OUT", "REDIR_APPEND"
                If Not hasName Then
                    Err.Raise vbObjectError + 1004, "clsParser.Parse", "Comando vazio antes do redirecionamento."
                End If
                If expectsTarget Then
                    Err.Raise vbObjectError + 1005, "clsParser.Parse", "Destino de redirecionamento ausente."
                End If
                Set redirection = New clsRedirection
                redirection.Operator = token.TokenType
                expectsTarget = True
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