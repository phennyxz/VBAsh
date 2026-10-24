VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Resolve(ByVal raw As String) As clsCommand
    Dim cmd As clsCommand
    Set cmd = New clsCommand
    cmd.Init Trim$(raw), raw
    Set Resolve = cmd
End Function
