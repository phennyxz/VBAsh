' @VBAshRevision clsBuiltinRegistry-r1
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
    Set builtin = New clsBuiltinBracketTest
    Register "[", builtin
End Sub

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsBuiltinRegistry-r1"
End Property

Private Sub Register(ByVal name As String, ByVal builtin As IBuiltin)
    Set mBuiltins(name) = builtin
End Sub

Public Sub RegisterBuiltin(ByVal commandName As String, ByVal builtin As IBuiltin, _
                           Optional ByVal replaceExisting As Boolean = False)
    Dim normalizedName As String

    normalizedName = Trim$(commandName)
    If Len(normalizedName) = 0 Or InStr(1, normalizedName, " ", vbBinaryCompare) > 0 Or _
       InStr(1, normalizedName, vbTab, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1320, "clsBuiltinRegistry.RegisterBuiltin", _
                  "Nome de builtin inválido."
    End If
    If builtin Is Nothing Then
        Err.Raise vbObjectError + 1321, "clsBuiltinRegistry.RegisterBuiltin", _
                  "A implementação do builtin é obrigatória."
    End If
    If mBuiltins.Exists(normalizedName) And Not replaceExisting Then
        Err.Raise vbObjectError + 1322, "clsBuiltinRegistry.RegisterBuiltin", _
                  "Builtin já registrado: " & normalizedName
    End If

    Set mBuiltins(normalizedName) = builtin
End Sub

Public Function HasBuiltin(ByVal commandName As String) As Boolean
    If Len(Trim$(commandName)) = 0 Then Exit Function
    HasBuiltin = mBuiltins.Exists(Trim$(commandName))
End Function

Public Function UnregisterBuiltin(ByVal commandName As String) As Boolean
    Dim normalizedName As String

    normalizedName = Trim$(commandName)
    If Len(normalizedName) = 0 Then Exit Function
    If Not mBuiltins.Exists(normalizedName) Then Exit Function

    mBuiltins.Remove normalizedName
    UnregisterBuiltin = True
End Function

Public Function Resolve(ByVal commandName As String) As IBuiltin
    If mBuiltins.Exists(commandName) Then
        Set Resolve = mBuiltins(commandName)
    End If
End Function
