VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFDs As Collection

Public Sub Initialize()
    Set mFDs = New Collection
    AddFD 0, CreateStream("stdin", "console")
    AddFD 1, CreateStream("stdout", "console")
    AddFD 2, CreateStream("stderr", "console")
End Sub

Private Function CreateStream(ByVal streamName As String, ByVal streamType As String) As clsStream
    Dim stream As clsStream
    Set stream = New clsStream
    stream.Initialize streamName, streamType
    Set CreateStream = stream
End Function

Public Sub AddFD(ByVal fd As Long, ByVal stream As clsStream)
    On Error Resume Next
    mFDs.Remove CStr(fd)
    On Error GoTo 0
    mFDs.Add stream, CStr(fd)
End Sub

Public Function GetFD(ByVal fd As Long) As clsStream
    On Error Resume Next
    Set GetFD = mFDs(CStr(fd))
    On Error GoTo 0
End Function

Public Sub WriteLine(ByVal text As String)
    Debug.Print text
End Sub

Public Sub Write(ByVal text As String)
    Debug.Print text;
End Sub
