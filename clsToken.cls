VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TokenType As String
Public Value As String

Public Sub Init(ByVal pType As String, ByVal pValue As String)
    TokenType = pType
    Value = pValue
End Sub
