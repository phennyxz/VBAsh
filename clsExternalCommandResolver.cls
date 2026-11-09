' @VBAshRevision clsExternalCommandResolver-r1
Option Explicit

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsExternalCommandResolver-r1"
End Property

Public Function ResolveExecutable(ByVal commandName As String, _
                                  ByVal sessionContext As clsSession) As String
    Dim fileSystem As Object
    Dim pathValue As String
    Dim pathEntries As Variant
    Dim extensions As Variant
    Dim entry As Variant
    Dim extension As Variant
    Dim candidate As String
    Dim hasExtension As Boolean
    Dim i As Long

    commandName = Trim$(commandName)
    If Len(commandName) = 0 Then Exit Function
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    If InStr(1, commandName, "\", vbBinaryCompare) > 0 Or _
       InStr(1, commandName, "/", vbBinaryCompare) > 0 Or _
       (Len(commandName) >= 2 And Mid$(commandName, 2, 1) = ":") Then
        candidate = commandName
        If Not fileSystem.FileExists(candidate) Then
            candidate = fileSystem.GetAbsolutePathName( _
                fileSystem.BuildPath(sessionContext.CurrentDirectory, commandName))
        End If
        If fileSystem.FileExists(candidate) Then ResolveExecutable = candidate
        Exit Function
    End If

    hasExtension = (InStrRev(commandName, ".") > 0)
    If hasExtension Then
        extensions = Array("")
    Else
        pathValue = Environ$("PATHEXT")
        If Len(pathValue) = 0 Then pathValue = ".COM;.EXE;.BAT;.CMD"
        extensions = Split(pathValue, ";")
    End If

    If sessionContext.Env.Exists("PATH") Then
        pathValue = sessionContext.GetEnv("PATH")
    Else
        pathValue = Environ$("PATH")
    End If
    If Len(pathValue) = 0 Then Exit Function
    pathEntries = Split(pathValue, ";")

    For Each entry In pathEntries
        If Len(Trim$(CStr(entry))) > 0 Then
            For Each extension In extensions
                candidate = fileSystem.BuildPath(Trim$(CStr(entry)), commandName & CStr(extension))
                If fileSystem.FileExists(candidate) Then
                    ResolveExecutable = candidate
                    Exit Function
                End If
            Next extension
        End If
    Next entry
End Function
