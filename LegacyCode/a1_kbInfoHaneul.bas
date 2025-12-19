Attribute VB_Name = "a1_kbInfoHaneul"
Option Explicit
'통합실행
Sub UpdateTableKB()

    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Tpl_Report_인포통계 시트를 참조
    Set ws = ThisWorkbook.Sheets("Tpl_Report_KB시세")
    
    ' tableAnalysis 테이블을 참조
    Set tbl = ws.ListObjects("tableValuation")
    
    ' 테이블에 데이터가 있을 때만 삭제
    If Not tbl.DataBodyRange Is Nothing Then
        ' 테이블에 데이터가 있을 경우 모든 행을 삭제
        tbl.DataBodyRange.Delete
    Else
        ' 데이터가 없을 경우: 테이블의 행을 직접 초기화 (빈 행 생성)
        tbl.Resize tbl.Range.Resize(1) ' 빈 테이블로 초기화
    End If
    
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("2-2")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    
    Call ReportKB
    Call TemplateKB
    Call Hyperlink_sheet("2-2", "Output_KB시세")
    Call AutoFitColumnsExceptA("2-2")
End Sub



'Output_KB시세 >> Report_KB시세
Sub ReportKB()
    Dim wsKB As Worksheet
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim lastRowKB As Long
    Dim i As Long
    Dim currentDate As String
    Dim insertCount As Long

    ' 시트 설정
    Set wsKB = ThisWorkbook.Sheets("Output_KB시세")
    Set wsData = ThisWorkbook.Sheets("Tpl_Report_KB시세")
    
    ' 테이블 존재 여부 확인
    On Error Resume Next
    Set tbl = wsData.ListObjects("tableValuation")
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Error: Table 'tableValuation' not found in the sheet 'Report_KB시세'.", vbCritical
        Exit Sub
    End If

    ' 마지막 행 찾기 (KB 시세 시트의)
    lastRowKB = wsKB.Cells(wsKB.Rows.Count, 1).End(xlUp).row

    ' KB 시세 시트에서 각 행을 순회하면서 값 복사
    insertCount = 0
    For i = 2 To lastRowKB ' Assuming the first row contains headers
        With wsKB
            Dim regNum As String
            Dim address As String
            Dim price As Variant
            Dim costDate As String
            
            regNum = .Cells(i, .columns("A").Column).value ' "등기부등본고유번호" column
            address = .Cells(i, .columns("C").Column).value ' "등기부등본주소" column
            price = .Cells(i, .columns("H").Column).value ' "일반가" column
            costDate = .Cells(i, .columns("K").Column).value ' "기준일자" column
        End With

        ' '일반가'가 있는 행에 대해서만 처리
        If Not IsEmpty(price) Then
            ' 새 행을 테이블에 추가
            With tbl.ListRows.Add
                ' 데이터 입력
                .Range.Cells(1, tbl.ListColumns("물건번호").Index).value = regNum
                .Range.Cells(1, tbl.ListColumns("담보물 주소").Index).value = address
                .Range.Cells(1, tbl.ListColumns("감정평가액").Index).value = price
                .Range.Cells(1, tbl.ListColumns("적용 감정가").Index).value = "KB시세조회"
                .Range.Cells(1, tbl.ListColumns("평가기관").Index).value = "KB부동산시세"
                .Range.Cells(1, tbl.ListColumns("감정평가기준일").Index).value = costDate
                ' 바탕색 흰색으로 설정
                .Range.Interior.Color = vbWhite
            End With
            insertCount = insertCount + 1
        End If
    Next i

End Sub


'Report_KB시세 >> 템플릿 작성
Sub TemplateKB()
    Dim wsTemp As Worksheet
    Dim wsReport As Worksheet
    Dim tbl As ListObject
    Dim areaRange As Range
    Dim destCell As Range
    Dim lastRow As Long
    Dim i As Long

    ' 1. 2-2 시트 생성 또는 초기화

    Dim wsName As String
    wsName = "2-2"
    
    ' "2-2" 시트가 있는지 확인하고, 없으면 생성
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    
    If wsTemp Is Nothing Then
        ' "2-2" 시트가 없으면 새로 생성
        Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTemp.Name = wsName
    End If


    ' 2. areaValuation 영역을 복사하여 2-2 B2에 붙여넣기
    Set wsReport = ThisWorkbook.Sheets("Tpl_report_area")
    Set areaRange = wsReport.Range("areaValuation")
    areaRange.Copy wsTemp.Range("B2")
    
    ' 복사된 영역의 행 개수
    Dim areaRowCount As Long
    areaRowCount = areaRange.Rows.Count
    
    ' 3. tableValuation 표 순환 및 값 입력
    Set tbl = ThisWorkbook.Sheets("Tpl_Report_KB시세").ListObjects("tableValuation")
    
    ' 열 번호 지정 (해당 표의 열 번호를 정확히 기입해야 함)
    Dim colAppraisalAmount As Long
    Dim colAppliedAppraisal As Long
    Dim colEvaluationAgency As Long
    Dim colAppraisalDate As Long
    
    colAppraisalAmount = tbl.ListColumns("감정평가액").Index
    colAppliedAppraisal = tbl.ListColumns("적용 감정가").Index
    colEvaluationAgency = tbl.ListColumns("평가기관").Index
    colAppraisalDate = tbl.ListColumns("감정평가기준일").Index

    lastRow = 1 ' 초기화
    
    For i = 1 To tbl.ListRows.Count
        ' 복사된 영역의 (2,5) 셀에 감정평가액 입력
        wsTemp.Cells(lastRow + 2, 6).value = tbl.ListRows(i).Range(1, colAppraisalAmount).value
        
        ' 복사된 영역의 (2,2) 셀에 적용 감정가 입력
        wsTemp.Cells(lastRow + 2, 3).value = tbl.ListRows(i).Range(1, colAppliedAppraisal).value
        
        ' 복사된 영역의 (3,2) 셀에 평가기관 입력
        wsTemp.Cells(lastRow + 3, 3).value = tbl.ListRows(i).Range(1, colEvaluationAgency).value
        
        ' 복사된 영역의 (3,5) 셀에 감정평가기준일 입력
        wsTemp.Cells(lastRow + 3, 6).value = tbl.ListRows(i).Range(1, colAppraisalDate).value
        
        ' 다음 행으로 복사된 영역 붙여넣기 (마지막 반복이 아닐 때만 복사)
        If i < tbl.ListRows.Count Then
            Set destCell = wsTemp.Cells(lastRow + areaRowCount + 2, 2)
            areaRange.Copy destCell
            lastRow = destCell.row - 1 ' 마지막 행 위치 갱신
        End If
    Next i
End Sub






