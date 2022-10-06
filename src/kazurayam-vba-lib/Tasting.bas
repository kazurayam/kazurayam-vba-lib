Attribute VB_Name = "Tasting"
Option Explicit

'Tasting --- ‚¢‚ë‚¢‚ë–¡Œ©‚·‚é

Sub tasteKzImportWorksheetFromWorkbook()
    'Arrange
    Dim wbSource As Workbook: Set wbSource = ActiveWorkbook
    Dim sourceSheetName As String: sourceSheetName = "Sheet1"
    Dim targetSheetName As String: targetSheetName = "work"
    'Act
    Call KzImportWorksheetFromWorkbook(wbSource, sourceSheetName, targetSheetName)
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
End Sub

