VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mSourceType As String
Private mBuffer As String

Public Sub Initialize(ByVal streamName As String, ByVal streamType As String)
    mName = streamName
    mSourceType = streamType
    mBuffer = ""
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get SourceType() As String
    SourceType = mSourceType
End Property

Public Function WriteText(ByVal text As String) As Long
    mBuffer = mBuffer & text
    WriteText = Len(text)
End Function

Public Function ReadText(Optional ByVal maxChars As Long = 32767) As String
    If Len(mBuffer) = 0 Then
        ReadText = ""
        Exit Function
    End If

    If Len(mBuffer) <= maxChars Then
        ReadText = mBuffer
        mBuffer = ""
    Else
        ReadText = Left$(mBuffer, maxChars)
        mBuffer = Mid$(mBuffer, maxChars + 1)
    End If
End Function

Public Sub Clear()
    mBuffer = ""
End Sub
