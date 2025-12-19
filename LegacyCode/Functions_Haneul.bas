Attribute VB_Name = "Functions_Haneul"
Option Explicit

' 시트가 이미 존재하는지 확인하는 함수
Function CheckSheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    CheckSheetExists = False
    For Each ws In wb.Sheets
        If ws.Name = sheetName Then
            CheckSheetExists = True
            Exit Function
        End If
    Next ws
End Function

'일자를 계산하는 함수
Function GetPreviousDate(baseDate As Date, period As String) As Date

    Dim previousDate As Date
    
    ' 조회기간에 따라 계산
    Select Case period
        Case "1년"
            previousDate = DateAdd("yyyy", -1, baseDate) ' 1년 전 날짜 계산
        Case "6개월"
            previousDate = DateAdd("m", -6, baseDate) ' 6개월 전 날짜 계산
        Case "3개월"
            previousDate = DateAdd("m", -3, baseDate) ' 3개월 전 날짜 계산
        Case Else
            MsgBox "조회기간은 '1년', '6개월', '3개월' 중 하나여야 합니다.", vbExclamation
            Exit Function
    End Select
    
    ' 결과 반환
    GetPreviousDate = previousDate

End Function

' 특정 헤더명을 기준으로 열 번호를 반환하는 함수 (B6셀부터 시작하는 헤더를 고려)
Function GetColumnNumber(ws As Worksheet, headerName As String, headerRow As Long) As Long
    Dim headerRange As Range
    Dim foundHeader As Range
    
    ' 헤더가 있는 행에서 헤더를 찾음
    Set headerRange = ws.Rows(headerRow)
    Set foundHeader = headerRange.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundHeader Is Nothing Then
        GetColumnNumber = foundHeader.Column
    Else
        MsgBox "헤더 '" & headerName & "'를 찾을 수 없습니다.", vbExclamation
        GetColumnNumber = -1 ' 헤더를 찾지 못했을 경우 -1 반환
    End If
End Function



