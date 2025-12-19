Attribute VB_Name = "commonHaneul"
'각 페이지 초기화
Sub ClearSheetsAndTables()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblName As String
    Dim targetTables As Variant
    Dim i As Long
    
    ' 첫 번째 작업: 특정 시트의 모든 값 지우기 및 배경색 초기화
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "dataAnalysis" Or ws.Name = "summaryAnalysis" Or ws.Name = "TransformedData" Then
            ' 시트의 모든 값을 지우고 배경색을 투명색으로 설정
            ws.Cells.Clear
            ws.Cells.Interior.ColorIndex = xlNone
        End If
    Next ws
    
    ' 두 번째 작업: 모든 시트의 특정 테이블 행 삭제
    targetTables = Array("tableValuation", "tableAuction", "tableAnalysis", "tableCases")
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            For i = LBound(targetTables) To UBound(targetTables)
                tblName = targetTables(i)
                If tbl.Name = tblName Then
                    ' 테이블의 모든 행 삭제
                    On Error Resume Next ' 에러가 발생해도 무시
                    tbl.DataBodyRange.Delete
                    On Error GoTo 0 ' 에러 무시 중지
                End If
            Next i
        Next tbl
    Next ws
    
    MsgBox "시트와 테이블의 정리 작업이 완료되었습니다!"
End Sub



'범위명과 셀이름을 입력받아, 셀 이름 위에 한 줄을 남기고 범위를 붙여넣는 함수
Sub CopyAndInsertRows(areaRange, targetcell)

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim rngSource As Range
    Dim rngDest As Range
    Dim insertRows As Long
    
    
    ' 소스 시트 및 대상 시트를 설정
    Set wsSource = ThisWorkbook.Sheets("Tpl_report_area")
    Set wsDest = ThisWorkbook.Sheets("Tpl_report")
    
    ' 소스 범위 및 대상 셀을 설정
    Set rngSource = wsSource.Range(areaRange)
    Set rngDest = wsDest.Range(targetcell)
    
    ' 삽입할 행의 개수를 계산
    insertRows = rngSource.Rows.Count
    
    ' 대상 셀 위에 소스 범위의 크기만큼 행을 삽입
    rngDest.EntireRow.Resize(insertRows + 1).Insert Shift:=xlDown  ' 1줄 추가로 더 삽입
    
    ' 행을 삽입한 후, 한 줄 아래로 이동하여 rngDest를 다시 설정 (한 줄 띄우기)
    Set rngDest = rngDest.offset(-insertRows - 1, 1) ' 1줄 아래로 이동
    
    ' 소스 범위를 복사하고 한 줄 띄운 후 붙여넣기
    rngSource.Copy
    rngDest.PasteSpecial Paste:=xlPasteAll
    
    ' 클립보드 비우기
    Application.CutCopyMode = False
End Sub

'시트명과 컬럼, 키워드를 입력받아,해당 시트의 컬럼에 특정 키워드로 시작하는 행을 삭제하는 함수
Sub DeleteRowsWithText(sheetName As String, columnName As String, keyword As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 지정한 시트를 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' 시트가 존재하지 않을 경우 종료
    If ws Is Nothing Then
        MsgBox "시트 '" & sheetName & "'를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' D열의 마지막 행을 찾음
    lastRow = ws.Cells(ws.Rows.Count, columnName).End(xlUp).row
    
    ' 아래에서 위로 반복하면서 지정된 키워드로 시작하는 텍스트를 찾아 해당 행을 삭제함.
    For i = lastRow To 1 Step -1
        If Left(ws.Cells(i, columnName).value, Len(keyword)) = keyword Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

'시트를 이름순서대로 정렬
Sub SortSheetsByName()
    Dim i As Integer, j As Integer
    Dim sheetCount As Integer
    Dim sheetNames() As String
    Dim tempName As String

    ' 현재 워크북의 시트 개수
    sheetCount = ThisWorkbook.Sheets.Count

    ' 시트 이름을 배열에 저장
    ReDim sheetNames(1 To sheetCount)
    For i = 1 To sheetCount
        sheetNames(i) = ThisWorkbook.Sheets(i).Name
    Next i

    ' 시트 이름을 알파벳 순으로 정렬 (버블 정렬 사용)
    For i = 1 To sheetCount - 1
        For j = i + 1 To sheetCount
            If sheetNames(i) > sheetNames(j) Then
                ' 이름 교환
                tempName = sheetNames(i)
                sheetNames(i) = sheetNames(j)
                sheetNames(j) = tempName
            End If
        Next j
    Next i

    ' 정렬된 순서대로 시트를 이동
    For i = 1 To sheetCount
        ThisWorkbook.Sheets(sheetNames(i)).Move Before:=ThisWorkbook.Sheets(i)
    Next i
End Sub

Sub HideOutputSheet(sheetName)
    Dim ws As Worksheet
    
    ' 모든 시트를 순회
    For Each ws In ThisWorkbook.Sheets
        ' 시트 이름이 "Output_"으로 시작하는지 확인
        If Left(ws.Name, 7) = sheetName Then
            ws.Visible = xlSheetHidden ' 시트를 숨김
        End If
    Next ws
End Sub

Sub DeleteAllRowsInTable(sheetName As String, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' 지정된 시트를 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 지정된 테이블을 설정
    Set tbl = ws.ListObjects(tableName)
    
    ' 테이블에 데이터가 있을 때만 삭제
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
End Sub

Sub DeleteColumnByHeader(sheetName As String, headerName As String)
    Dim ws As Worksheet
    Dim headerRow As Integer
    Dim lastCol As Integer
    Dim col As Integer
    
    ' 지정된 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 헤더가 위치한 행 설정 (6행)
    headerRow = 6
    
    ' 마지막 열 찾기
    lastCol = ws.Cells(headerRow, ws.columns.Count).End(xlToLeft).Column
    
    ' 각 열의 헤더 확인
    For col = 1 To lastCol
        If ws.Cells(headerRow, col).value = headerName Then
            ' 헤더가 일치하는 열 삭제
            ws.columns(col).Delete
            Exit For ' 일치하는 열을 찾았으면 더 이상 반복할 필요 없음
        End If
    Next col
End Sub


