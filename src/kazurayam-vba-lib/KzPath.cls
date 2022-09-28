VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KzPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' member variables
Private path As String

Private Sub Class_Initialize()
    path = ""
    Debug.Print "initialized"
End Sub

' default property returns the path As String
Property Get MyPath() As String
Attribute MyPath.VB_UserMemId = 0
    MyPath = path
End Propertyf

Property Let MyPath(path As String)
    If path = "" Then
        Err.Rains 10000, , "path is empty"
    End If
    path = path
End Property


