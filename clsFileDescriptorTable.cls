Option Explicit

Private mDescriptors As Object

Private Sub Class_Initialize()
    Set mDescriptors = CreateObject("Scripting.Dictionary")
End Sub

Public Sub SetDescriptor(ByVal descriptor As Long, ByVal stream As Object)
    Set mDescriptors(CStr(descriptor)) = stream
End Sub

Public Function GetDescriptor(ByVal descriptor As Long) As Object
    If mDescriptors.Exists(CStr(descriptor)) Then
        Set GetDescriptor = mDescriptors(CStr(descriptor))
    End If
End Function