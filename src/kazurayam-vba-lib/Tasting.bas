Attribute VB_Name = "Tasting"
Option Explicit

'Tasting --- いろいろ味見する

Sub tasteKzImportWorksheetFromWorkbook()
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "Sheet1"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call KzImportWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
End Sub

