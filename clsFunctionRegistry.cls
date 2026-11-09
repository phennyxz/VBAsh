' @VBAshRevision clsFunctionRegistry-r1
Option Explicit

Private mFunctions As Object

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsFunctionRegistry-r1"
End Property

Private Sub Class_Initialize()
    Set mFunctions = CreateObject("Scripting.Dictionary")
    mFunctions.CompareMode = vbBinaryCompare
End Sub

Public Sub RegisterFunction(ByVal functionName As String, ByVal functionBody As clsCommand, _
                            Optional ByVal replaceExisting As Boolean = False)
    Dim normalizedName As String

    normalizedName = Trim$(functionName)
    If Not IsValidFunctionName(normalizedName) Then
        Err.Raise vbObjectError + 1430, "clsFunctionRegistry.RegisterFunction", _
                  "Nome de função inválido: " & functionName
    End If
    If functionBody Is Nothing Then
        Err.Raise vbObjectError + 1431, "clsFunctionRegistry.RegisterFunction", _
                  "O corpo AST da função é obrigatório."
    End If
    If mFunctions.Exists(normalizedName) And Not replaceExisting Then
        Err.Raise vbObjectError + 1432, "clsFunctionRegistry.RegisterFunction", _
                  "Função já registrada: " & normalizedName
    End If

    Set mFunctions(normalizedName) = functionBody
End Sub

Public Function HasFunction(ByVal functionName As String) As Boolean
    If Len(Trim$(functionName)) = 0 Then Exit Function
    HasFunction = mFunctions.Exists(Trim$(functionName))
End Function

Public Function Resolve(ByVal functionName As String) As clsCommand
    functionName = Trim$(functionName)
    If mFunctions.Exists(functionName) Then
        Set Resolve = mFunctions(functionName)
    End If
End Function

Public Function RemoveFunction(ByVal functionName As String) As Boolean
    functionName = Trim$(functionName)
    If Len(functionName) = 0 Then Exit Function
    If Not mFunctions.Exists(functionName) Then Exit Function
    mFunctions.Remove functionName
    RemoveFunction = True
End Function

Public Property Get Count() As Long
    Count = mFunctions.Count
End Property

Private Function IsValidFunctionName(ByVal functionName As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(functionName) = 0 Then Exit Function
    For i = 1 To Len(functionName)
        ch = Mid$(functionName, i, 1)
        If i = 1 Then
            If Not (ch = "_" Or (ch >= "A" And ch <= "Z") Or _
                    (ch >= "a" And ch <= "z")) Then Exit Function
        Else
            If Not (ch = "_" Or (ch >= "A" And ch <= "Z") Or _
                    (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9")) Then
                Exit Function
            End If
        End If
    Next i
    IsValidFunctionName = True
End Function
