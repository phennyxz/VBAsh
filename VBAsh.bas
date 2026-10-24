Option Explicit

Public gSession As clsSession
Public gConsole As clsConsole
Public gFDTable As clsFileDescriptorTable

Public Sub VBAsh_Start()
    Set gSession = New clsSession
    Set gConsole = New clsConsole
    Set gFDTable = New clsFileDescriptorTable

    gConsole.Initialize

    Do
        Dim line As String
        line = InputBox(gSession.Prompt, "VBAsh")

        If StrPtr(line) = 0 Then
            ' Usuário clicou em Cancelar
            Exit Do
        End If

        If Len(Trim$(line)) = 0 Then
            GoTo ContinueLoop
        End If

        Dim result As String
        result = gSession.ProcessInput(line)

        If Len(result) > 0 Then
            gConsole.WriteLine result
        End If

        If LCase$(Trim$(line)) = "exit" Then
            Exit Do
        End If

ContinueLoop:
    Loop
End Sub