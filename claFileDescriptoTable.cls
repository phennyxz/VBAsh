Option Explicit

Public Name As String
Public SourceType As String
Public Buffer As String

Public Sub Initialize(ByVal streamName As String, ByVal streamType As String)
    Name = streamName
    SourceType = streamType
    Buffer = ""
End Sub

Public Function WriteText(ByVal text As String) As Long
    Buffer = Buffer & text
    WriteText = Len(text)
End Function

Public Function ReadText(Optional ByVal maxChars As Long = 32767) As String
    If Len(Buffer) = 0 Then
        ReadText = ""
        Exit Function
    End If

    If Len(Buffer) <= maxChars Then
        ReadText = Buffer
        Buffer = ""
    Else
        ReadText = Left$(Buffer, maxChars)
        Buffer = Mid$(Buffer, maxChars + 1)
    End If
End Function

Public Sub Clear()
    Buffer = ""
End Sub