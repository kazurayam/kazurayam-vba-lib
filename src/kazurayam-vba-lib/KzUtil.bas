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


Public Function KzVarTypeAsString(ByVal var As Variant) As String
    ' ˆø”var‚Ìtype‚ğ’²‚×‚Ä•Ï”‚ÌŒ^‚ğ¦‚·•¶š—ñi"Integer"‚È‚Çj‚ğ•Ô‚·
    Dim typeValue As Long: typeValue = VarType(var)
    Dim result As String: result = "unknown"
    If typeValue = 2 Then
        result = "Integer"
    ElseIf typeValue = 3 Then
        result = "Long"
    ElseIf typeValue = 5 Then
        result = "Double"
    ElseIf typeValue = 8 Then
        result = "String"
    ElseIf typeValue = 11 Then
        result = "Boolean"
    ElseIf typeValue = 7 Then
        result = "Date"
    ElseIf typeValue = 9 Then
        result = "Object"
    ElseIf typeValue = 0 Then
        result = "Variant"
    ElseIf typeValue = 8200 Then
        result = "String()"
    ElseIf typeValue = 8194 Then
        result = "Integer()"
    Else
        result = Str(typeValue)
    End If
    KzVarTypeAsString = result
End Function
