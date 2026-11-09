' @VBAshRevision clsAliasRegistry-r1
Option Explicit

Private mAliases As Object

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsAliasRegistry-r1"
End Property

Private Sub Class_Initialize()
    Set mAliases = CreateObject("Scripting.Dictionary")
    mAliases.CompareMode = vbBinaryCompare
End Sub

Public Sub RegisterAlias(ByVal aliasName As String, ByVal replacementText As String, _
                         Optional ByVal replaceExisting As Boolean = False)
    Dim normalizedName As String

    normalizedName = Trim$(aliasName)
    If Not IsValidAliasName(normalizedName) Then
        Err.Raise vbObjectError + 1400, "clsAliasRegistry.RegisterAlias", _
                  "Nome de alias inválido: " & aliasName
    End If
    If mAliases.Exists(normalizedName) And Not replaceExisting Then
        Err.Raise vbObjectError + 1401, "clsAliasRegistry.RegisterAlias", _
                  "Alias já registrado: " & normalizedName
    End If

    mAliases(normalizedName) = replacementText
End Sub

Public Function HasAlias(ByVal aliasName As String) As Boolean
    Dim normalizedName As String

    normalizedName = Trim$(aliasName)
    If Len(normalizedName) = 0 Then Exit Function
    HasAlias = mAliases.Exists(normalizedName)
End Function

Public Function ResolveAlias(ByVal aliasName As String) As String
    Dim normalizedName As String

    normalizedName = Trim$(aliasName)
    If mAliases.Exists(normalizedName) Then
        ResolveAlias = CStr(mAliases(normalizedName))
    End If
End Function

Public Function RemoveAlias(ByVal aliasName As String) As Boolean
    Dim normalizedName As String

    normalizedName = Trim$(aliasName)
    If Len(normalizedName) = 0 Then Exit Function
    If Not mAliases.Exists(normalizedName) Then Exit Function

    mAliases.Remove normalizedName
    RemoveAlias = True
End Function

Public Property Get Count() As Long
    Count = mAliases.Count
End Property

Private Function IsValidAliasName(ByVal aliasName As String) As Boolean
    Dim forbidden As Variant
    Dim item As Variant

    If Len(aliasName) = 0 Then Exit Function
    forbidden = Array(" ", vbTab, vbCr, vbLf, "=", ";", "|", "&", ">", "<", _
                      "(", ")", "'", """", "#")
    For Each item In forbidden
        If InStr(1, aliasName, CStr(item), vbBinaryCompare) > 0 Then Exit Function
    Next item

    IsValidAliasName = True
End Function
