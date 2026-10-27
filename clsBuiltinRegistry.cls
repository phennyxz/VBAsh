Option Explicit

Private mBuiltins As Object

Private Sub Class_Initialize()
    Dim builtin As IBuiltin

    Set mBuiltins = CreateObject("Scripting.Dictionary")
    mBuiltins.CompareMode = vbTextCompare

    Set builtin = New clsBuiltinEcho
    Register "echo", builtin
    Set builtin = New clsBuiltinPwd
    Register "pwd", builtin
    Set builtin = New clsBuiltinCd
    Register "cd", builtin
    Set builtin = New clsBuiltinExport
    Register "export", builtin
    Set builtin = New clsBuiltinSetEnv
    Register "setenv", builtin
    Set builtin = New clsBuiltinUnset
    Register "unset", builtin
    Set builtin = New clsBuiltinHelp
    Register "help", builtin
    Set builtin = New clsBuiltinExit
    Register "exit", builtin
    Set builtin = New clsBuiltinTest
    Register "test", builtin
    Set builtin = New clsBuiltinLs
    Register "ls", builtin
    Set builtin = New clsBuiltinMkdir
    Register "mkdir", builtin
    Set builtin = New clsBuiltinClear
    Register "clear", builtin
End Sub

Private Sub Register(ByVal name As String, ByVal builtin As IBuiltin)
    Set mBuiltins(name) = builtin
End Sub

Public Function Resolve(ByVal commandName As String) As IBuiltin
    If mBuiltins.Exists(commandName) Then
        Set Resolve = mBuiltins(commandName)
    End If
End Function
