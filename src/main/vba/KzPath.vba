Attribute VB_Name = "KzPath"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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


