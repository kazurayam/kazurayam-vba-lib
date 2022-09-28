Attribute VB_Name = "KzUtil"
Option Explicit

'KzUtil

' Clear Immediate Window
' calls Debug.Print many times to output blank lines
' so that the immediate window is wiped out
Public Sub KzCls()
    Dim i As Long
    For i = 0 To 20
        Debug.Print
    Next i
End Sub


