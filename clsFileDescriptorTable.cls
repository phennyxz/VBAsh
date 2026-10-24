VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileDescriptorTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFDs As Collection

Public Sub InitializeDefault(ByVal console As clsConsole)
    Set mFDs = New Collection
    SetFD 0, console.GetFD(0)
    SetFD 1, console.GetFD(1)
    SetFD 2, console.GetFD(2)
End Sub

Public Sub SetFD(ByVal fd As Long, ByVal stream As clsStream)
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

Public Sub CloseFD(ByVal fd As Long)
    On Error Resume Next
    mFDs.Remove CStr(fd)
    On Error GoTo 0
End Sub
