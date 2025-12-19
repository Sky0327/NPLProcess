Attribute VB_Name = "Module_Output"
Sub Output_불러오기(ByVal directoryName As String, ByVal fileName As String)

    Dim baseFolderPath As String
    Dim fullFolderPath As String
    Dim outputFilePath As String
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range

    ' Source 시트의 B4 셀에서 베이스 폴더 경로 가져오기
    baseFolderPath = ThisWorkbook.Sheets("Source").Range("B4").value
    reportName = ThisWorkbook.Sheets("Source").Range("B2").value

    ' 베이스 경로와 디렉토리명을 결합하여 최종 폴더 경로 생성
    fullFolderPath = baseFolderPath & "\" & "Temp" & "\" & directoryName

    ' 최종 파일 경로를 폴더 경로와 파일명으로 구성
    outputFilePath = fullFolderPath & "\" & fileName & "_" & reportName & ".xlsx"

    ' 파일 경로가 존재하는지 확인
    If Dir(outputFilePath) <> "" Then
        On Error GoTo ErrorHandler ' 에러 핸들링 추가

        ' 파일 열기 (백그라운드에서)
        Set wbSource = Workbooks.Open(outputFilePath, ReadOnly:=True)

        ' 파일의 모든 시트를 복사
        For Each ws In wbSource.Sheets
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count) ' 현재 워크북에 시트 복사
            
            ' 새로 붙여넣은 시트의 참조
            With ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ' 첫 번째 행의 서식을 Bold로 설정
                .Rows(1).Font.Bold = True
                
                ' 데이터가 있는 마지막 행과 마지막 열 찾기
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
                lastCol = .Cells(1, .columns.Count).End(xlToLeft).Column
                
                ' 데이터 영역에 테두리 적용
                Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
                dataRange.Borders.LineStyle = xlContinuous
            End With
        Next ws
        
        ' 원본 파일 닫기
        wbSource.Close False
        MsgBox "모든 시트가 성공적으로 복사되었습니다."

    Else
        ' 파일이 존재하지 않을 경우
        MsgBox "파일이 확인되지 않습니다. 올바른 경로와 파일명을 입력해주세요."
    End If

    Exit Sub

ErrorHandler:
    MsgBox "파일을 여는 중 오류가 발생했습니다."
    If Not wbSource Is Nothing Then
        wbSource.Close False
    End If

End Sub

Sub Output_대상여부Column_추가(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetCol As Long
    Dim dataRange As Range
    Dim rng As Range

    ' 입력받은 시트명을 가진 시트를 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' 시트가 존재하지 않으면 오류 메시지 출력 후 종료
    If ws Is Nothing Then
        MsgBox "해당 시트가 존재하지 않습니다: " & sheetName, vbExclamation
        Exit Sub
    End If

    ' 마지막 사용된 행 찾기 (1열 기준)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' 마지막 사용된 열 찾기
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column

    ' 대상여부 열을 추가 (가장 오른쪽에 추가)
    targetCol = lastCol + 1
    ws.Cells(1, targetCol).value = "대상여부"  ' 1행에 헤더 추가
    
    ' 헤더를 볼드체로 설정
    ws.Cells(1, targetCol).Font.Bold = True

    ' 데이터를 연속된 행까지 'V'로 채우기
    Set dataRange = ws.Range(ws.Cells(2, targetCol), ws.Cells(lastRow, targetCol)) ' 2행부터 마지막 행까지
    dataRange.value = "V"
    
    ' "V"가 입력된 셀들을 가운데 정렬
    dataRange.HorizontalAlignment = xlCenter
    
    ' "V"가 입력된 셀의 채우기 색상을 RGB(251, 226, 213)로 설정
    dataRange.Interior.Color = RGB(251, 226, 213)

End Sub

Sub Output_거리계산_setting(sheetName As String, headerText As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim targetcell As Range

    ' 입력받은 시트명을 가진 시트를 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' 시트가 존재하지 않으면 오류 메시지 출력 후 종료
    If ws Is Nothing Then
        MsgBox "해당 시트가 존재하지 않습니다: " & sheetName, vbExclamation
        Exit Sub
    End If

    ' 마지막 사용된 행과 열 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column

    ' 연속된 데이터 범위 설정 (A1부터 마지막 행과 열까지)
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' 데이터를 B6부터 시작하는 위치로 잘라내기 이동
    Set targetcell = ws.Cells(6, 2) ' B6 셀
    dataRange.Cut Destination:=targetcell

    ' A2 셀에 인수로 받은 문자열 입력
    ws.Range("A2").value = headerText

    ' A2 셀의 글씨를 볼드체로 하고, 글씨 크기를 20으로 설정
    With ws.Range("A2").Font
        .Bold = True
        .Size = 20
    End With

End Sub
Sub Output_불러오기_등본목록()
    Call Output_불러오기("등기부등본 기본정보", "Output_등기부등본기본정보")
    DoEvents
    Call AdjustColumnWidth("Output_등본목록", 8, 70, 17, 15, 70, 70)
    Call ApplyColorFormatting_grey("Output_등본목록", "A", "B")
    Call ApplyColorFormatting_vividpersimon("Output_등본목록", "C", "D", "E")
    Call ChangeFontSize("Output_등본목록")
    Call ApplyAllBorders("Output_등본목록")
End Sub
Sub Output_불러오기_등본조회()
    Call Output_불러오기("등기부등본", "Output_등본조회")
    DoEvents
    Call AdjustColumnWidth("Output_등본조회", 20, 15, 66, 10, 22, 36, 10, 19, 14, 19, 58, 12)
    Call ApplyColorFormatting_grey("Output_등본조회", "A", "B", "C")
    Call ApplyColorFormatting_vividpersimon("Output_등본조회", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    Call ChangeFontSize("Output_등본조회")
    Call ApplyAllBorders("Output_등본조회")
    Call AdjustColumnWidth("Output_등본조회(전체)", 20, 77, 36, 27, 115)
    Call ApplyColorFormatting_grey("Output_등본조회(전체)", "A", "B")
    Call ApplyColorFormatting_vividpersimon("Output_등본조회(전체)", "C", "D", "E")
    Call ChangeFontSize("Output_등본조회(전체)")
    Call ApplyAllBorders("Output_등본조회(전체)")
End Sub
Sub Output_불러오기_공시지가()
    Call Output_불러오기("공시지가", "Output_공시지가")
    DoEvents
    Call AdjustColumnWidth("Output_공시지가", 20, 15, 66, 16, 18, 18, 20, 20, 17, 12, 12, 18, 22, 20, 20, 22, 22, 19, 16.5)
    Call ApplyColorFormatting_grey("Output_공시지가", "A", "B", "C")
    Call ApplyColorFormatting_vividpersimon("Output_공시지가", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S")
    Call ChangeFontSize("Output_공시지가")
    Call ApplyAllBorders("Output_공시지가")
    Call AdjustColumnWidth("Output_공시지가(전체)", 66, 25, 66)
    Call ApplyColorFormatting_grey("Output_공시지가(전체)", "A")
    Call ApplyColorFormatting_vividpersimon("Output_공시지가(전체)", "B", "C")
    Call ChangeFontSize("Output_공시지가(전체)")
    Call ApplyAllBorders("Output_공시지가(전체)")
End Sub
Sub Output_불러오기_실거래가조회_국토교통부()
    Call Output_불러오기("실거래가조회_국토교통부", "Output_국토교통부_실거래가조회")
    DoEvents
    Call Output_대상여부Column_추가("Output_실거래가_국토")
    Call AdjustColumnWidth("Output_실거래가_국토", 3, 20, 15, 66, 5, 5, 21, 13, 11, 8.5, 8.5, 8.5, 8.5, 19, 8.5, 13, 8.5, 8.5, 8.5, 12, 5, 8.5, 8.5, 8.5, 8.5, 8.5)
    Call ApplyColorFormatting_grey("Output_실거래가_국토", "A", "B", "C")
    Call ApplyColorFormatting_vividpersimon("Output_실거래가_국토", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X")
    Call ApplyColorFormatting_orange("Output_실거래가_국토", "Y")
    Call ChangeFontSize("Output_실거래가_국토")
    Call ApplyAllBorders("Output_실거래가_국토")
    Call Output_거리계산_setting("Output_실거래가_국토", "실거래가 조회(아파트, 오피스텔, 연립/다세대)")
End Sub
Sub Output_불러오기_거리계산_국토교통부()
    Call Output_불러오기("거리계산_국토교통부", "Output_거리계산_국토교통부")
    DoEvents
    Call AddDealDateColumn
    Call AdjustColumnWidth("Output_거리_국토", 20, 15, 66, 21, 13, 11, 8.5, 38, 20, 8.5, 13, 8.5, 8.5, 8.5, 10, 12, 5, 8.5, 8.5, 8.5, 15.5)
    Call ApplyColorFormatting_grey("Output_거리_국토", "A", "B", "C")
    Call ApplyColorFormatting_vividpersimon("Output_거리_국토", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T")
    Call AddReportIncludeColumn("Output_거리_국토")
    Call ChangeFontSize("Output_거리_국토")
    Call ApplyAllBorders("Output_거리_국토")
End Sub
Sub Output_불러오기_실거래가조회_밸류맵()
    Call Output_불러오기("실거래가조회_밸류맵", "Output_밸류맵_실거래가조회")
    DoEvents
    Call Output_대상여부Column_추가("Output_실거래가_밸류맵")
    Call AdjustColumnWidth("Output_실거래가_밸류맵", 3, 66, 38, 8.5, 8.5, 8.5, 10, 10, 8.5, 16, 12, 12, 12, 12, 8.5)
    Call ApplyColorFormatting_grey("Output_실거래가_밸류맵", "A")
    Call ApplyColorFormatting_vividpersimon("Output_실거래가_밸류맵", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")
    Call ApplyColorFormatting_orange("Output_실거래가_밸류맵", "N")
    Call ChangeFontSize("Output_실거래가_밸류맵")
    Call ApplyAllBorders("Output_실거래가_밸류맵")
    Call Output_거리계산_setting("Output_실거래가_밸류맵", "실거래가 조회(토지, 상가, 단독/다가구 등)")
    Call SpeedUp
    Call UpdateOutputWithInputValues
    Call SpeedDown
End Sub
Sub Output_불러오기_거리계산_밸류맵()
    Call Output_불러오기("거리계산_밸류맵", "Output_거리계산_밸류맵")
    DoEvents
    Call AdjustColumnWidth("Output_거리_밸류맵", 66, 8.5, 38, 8.5, 8.5, 8.5, 10, 10, 8.5, 16, 12, 12, 12, 12, 15.5)
    Call ApplyColorFormatting_grey("Output_거리_밸류맵", "A")
    Call ApplyColorFormatting_vividpersimon("Output_거리_밸류맵", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    Call AddReportIncludeColumn("Output_거리_밸류맵")
    Call ChangeFontSize("Output_거리_밸류맵")
    Call ApplyAllBorders("Output_거리_밸류맵")
End Sub
Sub Output_불러오기_KB시세()
    Call Output_불러오기("KB시세", "Output_KB시세")
    Call FormatTableByOutputSheet("Output_KB시세")
    Call AdjustColumnWidth("Output_KB시세", 18.5, 15, 65, 30, 18, 0, 14.5, 7, 11, 11, 11, 8.5, 8.5, 8.5, 8.5)
End Sub
Sub Output_불러오기_법원경매()
    Call Output_불러오기("법원경매", "Output_법원경매")
    Call FormatTableByOutputSheet("Output_법원경매")
    Call AdjustColumnWidth("Output_법원경매", 18.5, 15, 65, 35, 28, 25, 15, 15, 5, 5)
End Sub
Sub Output_불러오기_인포통계()
    Call Output_불러오기("인포통계", "Output_인포통계")
    Call FormatTableByOutputSheet("Output_인포통계")
    Call AutoFitColumnsBySheetName("Output_인포통계")
End Sub
Sub Output_불러오기_인포통합()
    Call Output_불러오기("인포통합", "Output_인포통합")
    Call FormatTableByOutputSheet("Output_인포통합")
    Call AutoFitColumnsBySheetName("Output_인포통합")
End Sub

Sub Output_불러오기_인포사례상세()
    Call Output_불러오기("인포사례상세", "Output_인포사례상세")
    Call FormatTableByOutputSheet("Output_인포사례상세")
    Call AutoFitColumnsBySheetName("Output_인포사례상세")
    Call AdjustColumnWidth("Output_인포사례상세", 18, 15, 40, 20, 15, 51, 45, 50, 8.5, 8.5, 93, 12.5, 12.5, 63, 8.5, 10, 10, 10, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5)
End Sub


