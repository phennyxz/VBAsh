Option Explicit

Private mBuiltinRegistry As clsBuiltinRegistry

Private Sub Class_Initialize()
    Set mBuiltinRegistry = New clsBuiltinRegistry
End Sub

Public Function Resolve(ByVal commandName As String) As IBuiltin
    Set Resolve = mBuiltinRegistry.Resolve(commandName)
End Function