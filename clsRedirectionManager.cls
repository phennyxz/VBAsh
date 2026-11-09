' @VBAshRevision clsRedirectionManager-r1
Option Explicit

Private Const REDIRECTION_ERROR As Long = vbObjectError + 1270

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsRedirectionManager-r1"
End Property

' Applies redirections in lexical order. The caller owns the surrounding
' descriptor scope and must restore it after command execution.
Public Sub ApplyRedirections(ByVal table As clsFileDescriptorTable, _
                             ByVal scope As clsDescriptorScope, _
                             ByVal redirections As Collection, _
                             ByVal currentDirectory As String)
    Dim redirection As clsRedirection
    Dim descriptor As Long
    Dim targetPath As String
    Dim stream As IStream
    Dim fileStream As clsFileStream
    Dim nullStream As clsNullStream

    If table Is Nothing Then
        Err.Raise REDIRECTION_ERROR, "clsRedirectionManager", "Tabela de descritores ausente."
    End If
    If scope Is Nothing Then
        Err.Raise REDIRECTION_ERROR, "clsRedirectionManager", "Escopo de descritores ausente."
    End If
    If redirections Is Nothing Then Exit Sub

    For Each redirection In redirections
        Select Case redirection.Operator
            Case "REDIR_OUT", "REDIR_APPEND", "REDIR_ERR", "REDIR_ERR_APPEND", "REDIR_IN"
                If Len(redirection.Target) = 0 Then
                    Err.Raise REDIRECTION_ERROR, "clsRedirectionManager", _
                              "Destino de redirecionamento vazio."
                End If
                If redirection.Operator = "REDIR_IN" Then
                    descriptor = 0
                ElseIf Left$(redirection.Operator, 9) = "REDIR_ERR" Then
                    descriptor = 2
                Else
                    descriptor = 1
                End If

                If StrComp(redirection.Target, "/dev/null", vbTextCompare) = 0 Or _
                   StrComp(redirection.Target, "NUL", vbTextCompare) = 0 Then
                    Set nullStream = New clsNullStream
                    If descriptor = 0 Then nullStream.Initialize True, False
                    table.SetDescriptor descriptor, nullStream, True
                Else
                    targetPath = ResolvePath(currentDirectory, redirection.Target)
                    Set fileStream = New clsFileStream
                    If redirection.Operator = "REDIR_IN" Then
                        fileStream.OpenForRead targetPath
                    Else
                        fileStream.OpenForWrite targetPath, _
                            (redirection.Operator = "REDIR_APPEND" Or _
                             redirection.Operator = "REDIR_ERR_APPEND")
                    End If
                    Set stream = fileStream
                    table.SetDescriptor descriptor, stream, True
                End If

            Case "REDIR_ERR_TO_OUT"
                table.DuplicateDescriptor 1, 2

            Case "REDIR_DISCARD"
                table.CloseDescriptor 1

            Case "REDIR_ERR_DISCARD"
                table.CloseDescriptor 2

            Case Else
                Err.Raise REDIRECTION_ERROR, "clsRedirectionManager", _
                          "Operador de redirecionamento não suportado: " & redirection.Operator
        End Select
    Next redirection
End Sub

Private Function ResolvePath(ByVal currentDirectory As String, ByVal path As String) As String
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Len(path) >= 2 And Mid$(path, 2, 1) = ":" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 2) = "\\" Then
        ResolvePath = fileSystem.GetAbsolutePathName(path)
    ElseIf Left$(path, 1) = "\" Or Left$(path, 1) = "/" Then
        ResolvePath = fileSystem.GetAbsolutePathName(Left$(currentDirectory, 2) & path)
    Else
        ResolvePath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath(currentDirectory, path))
    End If
End Function
