VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Parse(ByVal tokens As Collection) As clsCommand
    Dim cmd As clsCommand
    Set cmd = New clsCommand

    If tokens Is Nothing Or tokens.Count = 0 Then
        cmd.Init "", ""
        Set Parse = cmd
        Exit Function
    End If

    Dim tk As clsToken
    Set tk = tokens(1)
    cmd.Init tk.Value, tk.Value

    Dim i As Long
    For i = 2 To tokens.Count
        Set tk = tokens(i)
        If tk.TokenType = "WORD" Then
            cmd.AddArg tk.Value
        ElseIf tk.TokenType = "EOF" Then
            Exit For
        End If
    Next i

    Set Parse = cmd
End Function
