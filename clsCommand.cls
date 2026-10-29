Option Explicit

Public Name As String
Public Args As Collection
Public Redirections As Collection
Public OriginalText As String
Public NodeType As String
Public Words As Collection
Public Children As Collection
Public SourceStart As Long
Public SourceEnd As Long

Private Sub Class_Initialize()
    Set Args = New Collection
    Set Redirections = New Collection
    Set Words = New Collection
    Set Children = New Collection
End Sub

Public Sub Init(ByVal pName As String, ByVal pOriginalText As String)
    Name = pName
    OriginalText = pOriginalText
    Set Args = New Collection
    Set Redirections = New Collection
    Set Words = New Collection
    Set Children = New Collection
    NodeType = "SIMPLE_COMMAND"
    SourceStart = 0
    SourceEnd = 0
End Sub

Public Sub InitAstNode(ByVal pNodeType As String, ByVal pSourceStart As Long, ByVal pSourceEnd As Long)
    Select Case pNodeType
        Case "SIMPLE_COMMAND", "SEQUENCE", "AND_IF", "OR_IF"
            NodeType = pNodeType
        Case Else
            Err.Raise vbObjectError + 1080, "clsCommand.InitAstNode", _
                "Tipo de nó AST inválido: " & pNodeType
    End Select

    SourceStart = pSourceStart
    SourceEnd = pSourceEnd
    Set Args = New Collection
    Set Redirections = New Collection
    Set Words = New Collection
    Set Children = New Collection
End Sub

Public Sub AddArg(ByVal arg As String)
    Args.Add arg
End Sub

Public Sub AddRedirection(ByVal redirection As clsRedirection)
    Redirections.Add redirection
End Sub

Public Sub AddWordToken(ByVal token As clsToken)
    If NodeType <> "SIMPLE_COMMAND" Then
        Err.Raise vbObjectError + 1081, "clsCommand.AddWordToken", _
            "Somente nós de comando simples podem conter palavras."
    End If
    Words.Add token
End Sub

Public Sub AddChild(ByVal child As clsCommand)
    If NodeType = "SIMPLE_COMMAND" Then
        Err.Raise vbObjectError + 1082, "clsCommand.AddChild", _
            "Nós de comando simples não podem conter filhos."
    End If
    Children.Add child
End Sub