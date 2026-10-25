Option Explicit

Public Buffer As String

Public Sub Initialize()
    Buffer = ""
End Sub

Public Sub Write(ByVal text As String)
    Buffer = Buffer & text
    If Len(text) > 0 Then MsgBox text, vbInformation, "VBAsh"
End Sub

Public Sub WriteLine(ByVal text As String)
    Buffer = Buffer & text & vbCrLf
    MsgBox text, vbInformation, "VBAsh"
End Sub