Attribute VB_Name = "Module_ForDeveloper"
Option Explicit

Sub UnhideAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub


Sub HideTplSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Tpl_", vbTextCompare) > 0 Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub


Sub HideDBSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "DB_", vbTextCompare) > 0 Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub

Sub HideSourceSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Source", vbTextCompare) > 0 Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub

