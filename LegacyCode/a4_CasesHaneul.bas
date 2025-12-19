Attribute VB_Name = "a4_CasesHaneul"
Option Explicit

Sub UpdateTableCases()

    DeleteAllRowsInTable "Tpl_Report_낙찰사례", "tableCases"
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("5-2")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행

    Call CopyValuesToTableCasesWithFormula
    Call ProcessCasesData
    Call Hyperlink_sheet("5-2", "Output_인포사례상세")

End Sub

'dataTable 채우기
Sub CopyValuesToTableCasesWithFormula()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tbl As ListObject
    Dim cleanedValue As String
    Dim tblLastRow As Long
    Dim tblLastColumn As Long

    ' 시트 설정
    Set wsSource = ThisWorkbook.Sheets("Output_인포사례상세")
    Set wsDest = ThisWorkbook.Sheets("Tpl_Report_낙찰사례")
    
    ' tableCases 표 설정
    Set tbl = wsDest.ListObjects("tableCases")
    
    ' 원본 시트의 마지막 행 찾기 (헤더 아래부터 끝까지)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    
    ' 기존 표의 마지막 행과 열 찾기
    tblLastRow = tbl.ListRows.Count
    tblLastColumn = tbl.ListColumns.Count
    
    ' 표가 비어있는 경우 첫 행부터 시작하도록 설정
    If tblLastRow = 0 Then
        tblLastRow = 1
    Else
        tblLastRow = tblLastRow + 1
    End If
    
    ' 대상 시트의 tableCases 표에 데이터 복사
    For i = 2 To lastRow ' Assuming row 1 contains headers
        ' D열의 값을 첫 번째 열에 복사
        tbl.ListRows.Add
        tbl.ListRows(tblLastRow).Range(1, 1).value = wsSource.Cells(i, 4).value
        
        ' C열의 값을 두 번째 열에 복사
        tbl.ListRows(tblLastRow).Range(1, 2).value = wsSource.Cells(i, 3).value
        
        ' P열의 값을 세 번째 열에 복사
        tbl.ListRows(tblLastRow).Range(1, 3).value = wsSource.Cells(i, 16).value
        
        ' E열의 값을 네 번째 열에 복사 (단위 제거)
        cleanedValue = Replace(wsSource.Cells(i, 5).value, "원", "")
        Debug.Print (cleanedValue)
        tbl.ListRows(tblLastRow).Range(1, 4).value = Val(Replace(cleanedValue, ",", ""))
        
        ' Q열의 값을 다섯 번째 열에 복사 (단위 제거)
        cleanedValue = Replace(wsSource.Cells(i, 17).value, "원", "")
        tbl.ListRows(tblLastRow).Range(1, 5).value = Val(Replace(cleanedValue, ",", ""))
        
        ' 6열에 수식 추가 (5열 / 4열)
        tbl.ListRows(tblLastRow).Range(1, 6).FormulaR1C1 = "=RC[-1]/RC[-2]"
        
        ' I열의 값을 일곱 번째 열에 복사
        tbl.ListRows(tblLastRow).Range(1, 7).value = wsSource.Cells(i, 9).value
        
        ' J열의 값을 여덟 번째 열에 복사
        tbl.ListRows(tblLastRow).Range(1, 8).value = wsSource.Cells(i, 10).value
        
        ' O열의 값을 아홉 번째 열에 복사
        tbl.ListRows(tblLastRow).Range(1, 9).value = wsSource.Cells(i, 15).value
        
        ' 다음 표의 행으로 이동
        tblLastRow = tblLastRow + 1
    Next i
    
End Sub


Sub ProcessCasesData()
    Dim wsReport As Worksheet
    Dim ws5_2 As Worksheet
    Dim tableRange As Range
    Dim headerRange As Range
    Dim wsName As String
    wsName = "5-2"
    
    ' 시트 설정
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_낙찰사례")
    

    
    ' "5-2" 시트 새로 생성
    Set ws5_2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws5_2.Name = wsName
    
    ' "tableCases" 표의 범위 가져오기
    On Error Resume Next
    Set tableRange = wsReport.ListObjects("tableCases").Range
    On Error GoTo 0
    
    If tableRange Is Nothing Then
        MsgBox "표 'tableCases'를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' '5-2' 시트의 B2 셀부터 붙여넣기
    tableRange.Copy Destination:=ws5_2.Range("B2")

    ' 새로 붙여넣은 데이터가 표로 생성된 경우, 이를 범위로 변환
    Dim pastedTable As ListObject
    On Error Resume Next
    Set pastedTable = ws5_2.ListObjects(1)
    On Error GoTo 0
    
    If Not pastedTable Is Nothing Then
        pastedTable.Unlist
    End If

    ' 데이터 영역의 배경을 하얀색으로 지정
    ws5_2.Range("B3").Resize(tableRange.Rows.Count - 1, tableRange.columns.Count).Interior.Color = RGB(255, 255, 255)
    
    ' 헤더 영역의 배경을 회색으로 지정
    Set headerRange = ws5_2.Range("B2").Resize(1, tableRange.columns.Count)
    headerRange.Interior.Color = RGB(242, 242, 242)
    
    '폰트 색을 짙은 회색으로 지정
    Dim ws As Worksheet
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets("5-2")
    
    ' 시트의 모든 셀의 폰트 색상을 RGB(128,128,128)로 설정
    ws.Cells.Font.Color = RGB(128, 128, 128)
    ws.Cells.Font.Size = 9
    Call AdjustColumnWidth("5-2", 3, 17, 63)
    ' E열과 F열: 숫자를 쉼표로 구분하여 정수 형식으로 표시
    ws.columns("E").NumberFormat = "#,##0"  ' 쉼표로 구분된 정수 형식
    ws.columns("F").NumberFormat = "#,##0"  ' 쉼표로 구분된 정수 형식

    ' G열: 백분율 형식으로 표시
    ws.columns("G").NumberFormat = "0.00%"  ' 소수점 둘째 자리까지 백분율 형식
    
    
End Sub



