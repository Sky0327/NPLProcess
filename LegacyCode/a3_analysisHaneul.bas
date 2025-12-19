Attribute VB_Name = "a3_analysisHaneul"
'Output_인포통계 -> dataAnalysis(표 안의 표를 언피봇) -> summaryAnalysis(선택된 데이터 정리) -> tableAnalysis(유저 중간화면) -> tempAnalysis

Sub UpdateAnalysisTable()
    Dim SheetExists As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Tpl_Report_인포통계 시트를 참조
    Set ws = ThisWorkbook.Sheets("Tpl_Report_인포통계")
    
    ' tableAnalysis 테이블을 참조
    Set tbl = ws.ListObjects("tableAnalysis")
    
    ' 테이블에 데이터가 있을 때만 삭제
    If Not tbl.DataBodyRange Is Nothing Then
        ' 테이블에 데이터가 있을 경우 모든 행을 삭제
        tbl.DataBodyRange.Delete
    Else
        ' 데이터가 없을 경우: 테이블의 행을 직접 초기화 (빈 행 생성)
        tbl.Resize tbl.Range.Resize(1) ' 빈 테이블로 초기화
    End If
    
    
    SheetExists = Check_Sheet("dataAnalysis")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    
    Dim SheetExists2 As Boolean
    SheetExists2 = Check_Sheet("summaryAnalysis")
    If SheetExists2 = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    
    Dim SheetExists3 As Boolean
    SheetExists3 = Check_Sheet("5-1")
    If SheetExists3 = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    
 Call transformAnalysis
 Call ShowAuctionPeriodForCounts
 Call HighlightRowsByPeriod
 Call summaryAnalysis
 Call CopyToTableAnalysis
 Call CreateTempAnalysis
 Call ImportImagesFromFolder
 Call Hyperlink_sheet("5-1", "Output_인포통계")
 Call HideSheetByName("dataAnalysis")
 Call HideSheetByName("summaryAnalysis")

End Sub



'인포통계 Output >> dataAnalysis
Sub transformAnalysis()

    ' 원본 데이터 시트와 결과 시트 정의
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim wsName As String
    wsName = "dataAnalysis"
    
    ' 원본 시트 설정
    Set wsSrc = ThisWorkbook.Sheets("Output_인포통계") '하드코딩
    DeleteRowsWithText "Output_인포통계", "D", "조회"
    
    ' 결과 시트 확인 및 생성
    On Error Resume Next
    Set wsDst = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    
    ' "dataAnalysis" 시트가 없으면 생성
    If wsDst Is Nothing Then
        Set wsDst = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDst.Name = wsName
    End If
    
    ' 결과 시트 초기화
    wsDst.Cells.Clear
    
    ' 헤더 설정
    wsDst.Cells(1, 1).value = "고유번호"
    wsDst.Cells(1, 2).value = "등기부등본주소"
    wsDst.Cells(1, 3).value = "필터"
    wsDst.Cells(1, 4).value = "지역(대)"
    wsDst.Cells(1, 5).value = "지역(중)"
    wsDst.Cells(1, 6).value = "지역(소)"
    
    ' 낙찰 관련 헤더
    Dim periods As Variant
    periods = Array("1년", "6개월", "3개월")
    
    Dim scope As Variant
    scope = Array("대", "중", "소")
    
    Dim col As Integer
    col = 7
    
    Dim i As Integer, j As Integer
    For j = LBound(scope) To UBound(scope)
        For i = LBound(periods) To UBound(periods)
            wsDst.Cells(1, col).value = "낙찰가율(" & scope(j) & "_" & periods(i) & ")"
            wsDst.Cells(1, col + 1).value = "낙찰률 평균(" & scope(j) & "_" & periods(i) & ")"
            wsDst.Cells(1, col + 2).value = "낙찰건수(" & scope(j) & "_" & periods(i) & ")"
            col = col + 3
        Next i
    Next j

    ' 데이터 변환 작업
    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    
    Dim srcRow As Long
    Dim dstRow As Long
    dstRow = 2
    
    ' 데이터 초기값 변수 선언
    Dim uniqueId As String
    Dim regAddress As String
    Dim filterValue As String
    Dim regionLarge As String
    Dim regionMiddle As String
    Dim regionSmall As String
    
    ' 데이터 스캔
    For srcRow = 2 To lastRow Step 5 ' 5줄씩 읽기
        uniqueId = wsSrc.Cells(srcRow, 1).value
        regAddress = wsSrc.Cells(srcRow, 3).value
        filterValue = wsSrc.Cells(srcRow, 4).value
        regionLarge = wsSrc.Cells(srcRow, 6).value ' 지역(대)
        regionMiddle = wsSrc.Cells(srcRow, 9).value ' 지역(중)
        regionSmall = wsSrc.Cells(srcRow, 12).value ' 지역(소)
        
        ' 결과 시트에 고유 정보 입력
        wsDst.Cells(dstRow, 1).value = uniqueId
        wsDst.Cells(dstRow, 2).value = regAddress
        wsDst.Cells(dstRow, 3).value = filterValue
        wsDst.Cells(dstRow, 4).value = regionLarge
        wsDst.Cells(dstRow, 5).value = regionMiddle
        wsDst.Cells(dstRow, 6).value = regionSmall
        
        col = 7
        
        ' 기간별, 구분별 데이터 입력
        Dim offset As Integer
        For j = 0 To 2 ' 대, 중, 소
            For i = 0 To 2 ' 1년, 6개월, 3개월 (역순)
                offset = srcRow + 2 + i ' 기준점에서 3줄씩 떨어진 곳에 데이터 있음 (역순)
                
                wsDst.Cells(dstRow, col).value = wsSrc.Cells(offset, 6 + (j * 3)).value ' 낙찰가율
                wsDst.Cells(dstRow, col + 1).value = wsSrc.Cells(offset, 7 + (j * 3)).value ' 낙찰률 평균
                wsDst.Cells(dstRow, col + 2).value = wsSrc.Cells(offset, 8 + (j * 3)).value ' 낙찰건수
                col = col + 3
            Next i
        Next j
        
        dstRow = dstRow + 1
    Next srcRow

End Sub


'인포통계 dataAnalysis 시트에서 낙찰가율 선택하여 노란색 음영표시
Sub ShowAuctionPeriodForCounts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dataAnalysis")
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long, col As Long
    Dim auctionRateCol As Long
    
    ' 마지막 행과 마지막 열 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column
    
    Dim periodInfo As String
    Dim found As Boolean
    
    ' 데이터 행 탐색 (2행부터 시작)
    For row = 2 To lastRow
        found = False
        ' 가장 우측 열부터 탐색
        For col = lastCol To 1 Step -1
            If ws.Cells(1, col).value Like "낙찰건수(*" Then
                ' 해당 '낙찰건수'와 연결된 '낙찰가율'의 열 찾기
                auctionRateCol = col - 2 ' 낙찰가율이 낙찰건수 열의 2열 앞에 있다고 가정
                
                ' 낙찰가율이 1 이하인지 확인
                If ws.Cells(row, auctionRateCol).value <= 1 Then
                    If ws.Cells(row, col).value >= 5 Then
                        periodInfo = ws.Cells(1, col).value
                        ' 낙찰가율이 1(100%) 이상인 경우 건너뜀
                        If ws.Cells(row, auctionRateCol).value >= 1 Then
                            GoTo ContinueSearch
                        End If
                        
                        ' 노란색 음영 추가
                        ws.Cells(row, col).Interior.Color = RGB(255, 255, 0)
                        found = True
                        
                        Exit For ' 첫 번째로 찾은 항목에서 중단
                    End If
                End If
            End If
ContinueSearch:
        Next col
        
        ' '낙찰건수'가 5 이상인 항목이 없을 경우
        If Not found Then
            MsgBox "고유번호: " & ws.Cells(row, 1).value & "에서 낙찰건수 5 이상인 데이터가 없습니다.", vbExclamation
        End If
    Next row
End Sub




'선택된 낙찰건수에 상응하는 기간에 주황색 하이라이트.
Sub HighlightRowsByPeriod()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim header As Range
    Dim targetPeriod As String
    
    ' dataAnalysis 시트 설정
    Set ws = ThisWorkbook.Sheets("dataAnalysis")
    
    ' 데이터 범위의 마지막 행과 마지막 열 구하기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column
    
    ' 헤더 범위 지정
    Set header = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    ' 각 행을 순회하면서 작업 수행
    For i = 2 To lastRow
        ' 각 열을 순회하면서 노란색 셀 찾기
        targetPeriod = ""
        
        For j = 1 To lastCol
            If ws.Cells(i, j).Interior.Color = RGB(255, 255, 0) Then ' 노란색으로 색칠된 셀 발견
                ' 해당 헤더에서 1년/6개월/3개월 여부 추출
                If InStr(header.Cells(1, j).value, "1년") > 0 Then
                    targetPeriod = "1년"
                ElseIf InStr(header.Cells(1, j).value, "6개월") > 0 Then
                    targetPeriod = "6개월"
                ElseIf InStr(header.Cells(1, j).value, "3개월") > 0 Then
                    targetPeriod = "3개월"
                End If
                Exit For ' 찾았으면 반복문 종료
            End If
        Next j
        
        ' 주어진 기간에 해당하는 셀들을 주황색으로 칠하기 (노란색 셀은 제외)
        If targetPeriod <> "" Then
            For j = 1 To lastCol
                If InStr(header.Cells(1, j).value, targetPeriod) > 0 Then
                    ' 현재 셀이 노란색이 아닌 경우에만 주황색으로 칠하기
                    If ws.Cells(i, j).Interior.Color <> RGB(255, 255, 0) Then
                        ws.Cells(i, j).Interior.Color = RGB(255, 165, 0) ' 주황색
                    End If
                End If
            Next j
        End If
    Next i
End Sub



'중간산출물인 Data로 변환
Sub summaryAnalysis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dataAnalysis")
    
    ' 기존 정리된표 시트가 있으면 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("summaryAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Dim outputWs As Worksheet
    Set outputWs = ThisWorkbook.Sheets.Add
    outputWs.Name = "summaryAnalysis"
    
    ' 헤더 작성
    outputWs.Cells(1, 1).value = "담보물주소"
    outputWs.Cells(1, 2).value = "출처"
    outputWs.Cells(1, 3).value = "용도"
    outputWs.Cells(1, 4).value = "적용율"
    outputWs.Cells(1, 5).value = "기간"
    outputWs.Cells(1, 6).value = "소재지1"
    outputWs.Cells(1, 7).value = "소재지1_낙찰가율(%)"
    outputWs.Cells(1, 8).value = "소재지1_낙찰건수"
    outputWs.Cells(1, 9).value = "소재지2"
    outputWs.Cells(1, 10).value = "소재지2_낙찰가율(%)"
    outputWs.Cells(1, 11).value = "소재지2_낙찰건수"
    outputWs.Cells(1, 12).value = "소재지3"
    outputWs.Cells(1, 13).value = "소재지3_낙찰가율(%)"
    outputWs.Cells(1, 14).value = "소재지3_낙찰건수"
    
    Dim i As Long, j As Long, outRow As Long
    outRow = 2
    
    ' 데이터 분석 시트의 각 행에 대해 반복
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        Dim 담보물주소 As String, 출처 As String, 용도 As String
        Dim 적용율 As Double, 기간 As String
        Dim 소재지1 As String, 소재지2 As String, 소재지3 As String
        Dim 소재지1_낙찰가율 As Double, 소재지2_낙찰가율 As Double, 소재지3_낙찰가율 As Double
        Dim 소재지1_낙찰건수 As Variant, 소재지2_낙찰건수 As Variant, 소재지3_낙찰건수 As Variant
        
        ' 기본 정보 설정
        담보물주소 = ws.Cells(i, 2).value
        출처 = "인포케어"
        용도 = Split(ws.Cells(i, 3).value, "_")(UBound(Split(ws.Cells(i, 3).value, "_")))
        소재지1 = ws.Cells(i, 4).value
        소재지2 = ws.Cells(i, 5).value
        소재지3 = ws.Cells(i, 6).value
        
        ' 초기화
        적용율 = 0
        기간 = ""
        소재지1_낙찰가율 = 0
        소재지1_낙찰건수 = ""
        소재지2_낙찰가율 = 0
        소재지2_낙찰건수 = ""
        소재지3_낙찰가율 = 0
        소재지3_낙찰건수 = ""
        
        ' 낙찰가율 및 낙찰건수 추출을 위한 색상 확인
        For j = 7 To 33
            If ws.Cells(i, j).Interior.Color = RGB(255, 255, 0) Then ' 노란색 셀
                적용율 = ws.Cells(i, j - 2).value
                기간 = Split(ws.Cells(1, j).value, "(")(1)
                
                기간 = Replace(기간, ")", "")
                기간 = Replace(기간, "중_", "") ' 중_ 제거
                기간 = Replace(기간, "대_", "") ' 대_ 제거
                기간 = Replace(기간, "소_", "") ' 소_ 제거
                
                ' 낙찰건수 추출 로직 추가
                Dim headerText As String
                headerText = ws.Cells(1, j).value
                
                If InStr(headerText, "대_") > 0 Then
                    If InStr(headerText, "낙찰건수") > 0 Then
                        소재지1_낙찰건수 = ws.Cells(i, j).value
                    End If
                ElseIf InStr(headerText, "중_") > 0 Then
                    If InStr(headerText, "낙찰건수") > 0 Then
                        소재지2_낙찰건수 = ws.Cells(i, j).value
                    End If
                ElseIf InStr(headerText, "소_") > 0 Then
                    If InStr(headerText, "낙찰건수") > 0 Then
                        소재지3_낙찰건수 = ws.Cells(i, j).value
                    End If
                End If
                
            ElseIf ws.Cells(i, j).Interior.Color = RGB(255, 165, 0) Then ' 주황색 셀
                headerText = ws.Cells(1, j).value
                
                If InStr(headerText, "대_") > 0 Then
                    If InStr(headerText, "낙찰가율") > 0 Then
                        소재지1_낙찰가율 = ws.Cells(i, j).value
                    ElseIf InStr(headerText, "낙찰건수") > 0 Then
                        소재지1_낙찰건수 = ws.Cells(i, j).value
                    End If
                ElseIf InStr(headerText, "중_") > 0 Then
                    If InStr(headerText, "낙찰가율") > 0 Then
                        소재지2_낙찰가율 = ws.Cells(i, j).value
                    ElseIf InStr(headerText, "낙찰건수") > 0 Then
                        소재지2_낙찰건수 = ws.Cells(i, j).value
                    End If
                ElseIf InStr(headerText, "소_") > 0 Then
                    If InStr(headerText, "낙찰가율") > 0 Then
                        소재지3_낙찰가율 = ws.Cells(i, j).value
                    ElseIf InStr(headerText, "낙찰건수") > 0 Then
                        소재지3_낙찰건수 = ws.Cells(i, j).value
                    End If
                End If
            End If
        Next j

        
        ' 정리된 표에 작성
        outputWs.Cells(outRow, 1).value = 담보물주소
        outputWs.Cells(outRow, 2).value = 출처
        outputWs.Cells(outRow, 3).value = 용도
        outputWs.Cells(outRow, 4).value = 적용율
        outputWs.Cells(outRow, 5).value = 기간
        outputWs.Cells(outRow, 6).value = 소재지1
        outputWs.Cells(outRow, 7).value = 소재지1_낙찰가율
        outputWs.Cells(outRow, 8).value = 소재지1_낙찰건수
        outputWs.Cells(outRow, 9).value = 소재지2
        outputWs.Cells(outRow, 10).value = 소재지2_낙찰가율
        outputWs.Cells(outRow, 11).value = 소재지2_낙찰건수
        outputWs.Cells(outRow, 12).value = 소재지3
        outputWs.Cells(outRow, 13).value = 소재지3_낙찰가율
        outputWs.Cells(outRow, 14).value = 소재지3_낙찰건수
        
        outRow = outRow + 1
    Next i

End Sub

'Tpl_Report_인포통계' 채우기
Sub CopyToTableAnalysis()
    Dim wsSource As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim copyRange As Range
    Dim tblLastRow As Long
    Dim tblLastColumn As Long
    Dim destRange As Range
    
    ' Set source worksheet
    Set wsSource = ThisWorkbook.Sheets("summaryAnalysis")
    
    ' Find the last used row in the source sheet (starting from row 1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    Debug.Print lastRow
    
    ' Set the table
    Set tbl = ThisWorkbook.Sheets("Tpl_Report_인포통계").ListObjects("tableAnalysis")
    
    ' Set the range to copy (excluding headers, from A2 to the last row and last column)
    Set copyRange = wsSource.Range("A2:" & wsSource.Cells(lastRow, wsSource.columns.Count).End(xlToLeft).address)
    
    ' Find the number of data rows and columns in the table
    tblLastRow = tbl.ListRows.Count
    tblLastColumn = tbl.ListColumns.Count
    Debug.Print tblLastRow
    Debug.Print tblLastColumn
    
    ' Check if the table is empty
    If tbl.DataBodyRange Is Nothing Then
        ' Resize the table to include new data rows
        tbl.Resize tbl.Range.Resize(1 + copyRange.Rows.Count, tblLastColumn)
        
        ' Set the destination range directly below the header
        Set destRange = tbl.Range.offset(1, 0).Resize(copyRange.Rows.Count, tblLastColumn)
    Else
        ' Resize the table to include new data rows
        tbl.Resize tbl.Range.Resize(1 + tblLastRow + copyRange.Rows.Count, tblLastColumn)
        
        ' Set the destination range to the new rows in DataBodyRange
        Set destRange = tbl.DataBodyRange.Rows(tblLastRow + 1).Resize(copyRange.Rows.Count)
    End If
    
    ' Copy the data to the destination range
    destRange.value = copyRange.value
    
    'MsgBox "Data has been successfully added to the tableAnalysis table!"
End Sub





'tableAnalysis -> tempAnalysis
Sub CreateTempAnalysis()

    Dim wsSource As Worksheet
    Dim wsTemp As Worksheet
    Dim tblAnalysis As ListObject
    Dim rngAreaInfo As Range
    Dim cell As Range
    Dim newSheetName As String
    Dim currentRow As Long
    Dim wsExists As Boolean
    Dim 출처 As String
    Dim 기간 As String
    Dim 소재지1 As String
    Dim 소재지2 As String
    Dim 소재지3 As String
    Dim 용도 As String
    Dim 적용율 As Double
    Dim 소재지1_낙찰가율 As Double
    Dim 소재지1_낙찰건수 As Double
    Dim 소재지2_낙찰가율 As Double
    Dim 소재지2_낙찰건수 As Double
    Dim 소재지3_낙찰가율 As Double
    Dim 소재지3_낙찰건수 As Double

    newSheetName = "5-1"

    ' tableAnalysis 표가 있는 시트와 범위 설정
    Set wsSource = ThisWorkbook.Sheets("Tpl_Report_인포통계")
    Set tblAnalysis = wsSource.ListObjects("tableAnalysis")

    ' areaInfocareAnalysis 범위 설정 (범위가 존재하는지 확인)
    On Error Resume Next
    Set rngAreaInfo = ThisWorkbook.Sheets("Tpl_Report_area").Range("areaInfocareAnalysis")
    On Error GoTo 0
    If rngAreaInfo Is Nothing Then
        MsgBox "'areaInfocareAnalysis' 범위를 찾을 수 없습니다."
        Exit Sub
    End If

    ' 기존 tempAnalysis 시트가 있으면 삭제하고 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    wsExists = Not ThisWorkbook.Sheets(newSheetName) Is Nothing
    If wsExists Then
        ThisWorkbook.Sheets(newSheetName).Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = newSheetName

    ' 첫 번째 표는 B2부터 시작
    currentRow = 2

    ' 데이터를 참조하여 변수에 할당
    For Each cell In tblAnalysis.ListColumns("출처").DataBodyRange
        출처 = cell.value
        기간 = cell.offset(0, tblAnalysis.ListColumns("조회기간").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지1 = cell.offset(0, tblAnalysis.ListColumns("지역1").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지2 = cell.offset(0, tblAnalysis.ListColumns("지역2").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지3 = cell.offset(0, tblAnalysis.ListColumns("지역3").Index - tblAnalysis.ListColumns("출처").Index).value
        용도 = cell.offset(0, tblAnalysis.ListColumns("구분").Index - tblAnalysis.ListColumns("출처").Index).value
        적용율 = cell.offset(0, tblAnalysis.ListColumns("적용율").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지1_낙찰가율 = cell.offset(0, tblAnalysis.ListColumns("낙찰가율1").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지1_낙찰건수 = cell.offset(0, tblAnalysis.ListColumns("낙찰건수1").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지2_낙찰가율 = cell.offset(0, tblAnalysis.ListColumns("낙찰가율2").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지2_낙찰건수 = cell.offset(0, tblAnalysis.ListColumns("낙찰건수2").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지3_낙찰가율 = cell.offset(0, tblAnalysis.ListColumns("낙찰가율3").Index - tblAnalysis.ListColumns("출처").Index).value
        소재지3_낙찰건수 = cell.offset(0, tblAnalysis.ListColumns("낙찰건수3").Index - tblAnalysis.ListColumns("출처").Index).value

        ' areaInfocareAnalysis 범위를 tempAnalysis 시트에 복사
        rngAreaInfo.Copy wsTemp.Cells(currentRow, 2)

        ' 데이터를 지정된 위치에 입력
        wsTemp.Cells(currentRow, 4).value = 출처
        wsTemp.Cells(currentRow + 1, 4).value = 기간
        wsTemp.Cells(currentRow + 2, 4).value = 소재지1
        wsTemp.Cells(currentRow + 2, 6).value = 소재지2
        wsTemp.Cells(currentRow + 2, 8).value = 소재지3
        wsTemp.Cells(currentRow + 4, 2).value = 용도
        wsTemp.Cells(currentRow + 4, 3).value = 적용율
        wsTemp.Cells(currentRow + 4, 4).value = 소재지1_낙찰가율
        wsTemp.Cells(currentRow + 4, 5).value = 소재지1_낙찰건수
        wsTemp.Cells(currentRow + 4, 6).value = 소재지2_낙찰가율
        wsTemp.Cells(currentRow + 4, 7).value = 소재지2_낙찰건수
        wsTemp.Cells(currentRow + 4, 8).value = 소재지3_낙찰가율
        wsTemp.Cells(currentRow + 4, 9).value = 소재지3_낙찰건수

        ' 다음 표를 위해 한 줄 띄우고 아래로 이동
        currentRow = currentRow + rngAreaInfo.Rows.Count + 1
    Next cell

    ' 클립보드를 비웁니다.
    Application.CutCopyMode = False


End Sub

'이미지 불러오기
Sub ImportImagesFromFolder()
    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim imgTop As Double
    Dim imgLeft As Double
    Dim maxCols As Integer
    Dim colCount As Integer
    Dim shp As Shape
    Dim scaleFactor As Double
    
    ' 크기 축소 비율 설정
    scaleFactor = 1
    
    ' "Source" 시트에서 경로 가져오기
    Set wsSource = ThisWorkbook.Sheets("Source")
    folderPath = wsSource.Range("B4").value
    
    ' 경로에 "Temp\인포통계" 추가
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & "Temp\인포통계\"
    
    ' 시트 설정 (이미지를 불러올 시트를 지정)
    Set ws = ThisWorkbook.Sheets("5-1") ' 원하는 시트 이름으로 변경
    
    ' 이미지 배치 설정
    imgTop = ws.Range("K2").Top    ' 시작 위치의 상단 여백
    imgLeft = ws.Range("K2").Left  ' 시작 위치의 좌측 여백
    maxCols = 5    ' 한 줄에 표시할 최대 이미지 개수
    
    colCount = 0   ' 현재 열의 이미지 개수 초기화
    
    ' 첫 번째 파일 가져오기 (jpg, png, gif 확장자에 대해)
    fileName = Dir(folderPath & "*.*")
    
    Do While fileName <> ""
        If LCase(Right(fileName, 4)) = ".jpg" Or LCase(Right(fileName, 4)) = ".png" Or LCase(Right(fileName, 4)) = ".gif" Then
            ' 이미지 삽입 (원본 크기로 삽입)
            Set shp = ws.Shapes.AddPicture(folderPath & fileName, _
                msoFalse, msoCTrue, imgLeft, imgTop, -1, -1) ' -1을 사용하여 원본 크기로 삽입
            
            ' 이미지 크기를 30%로 축소
            shp.LockAspectRatio = msoTrue ' 가로세로 비율 유지
            shp.Width = shp.Width * scaleFactor
            shp.Height = shp.Height * scaleFactor
            
            ' 다음 이미지의 위치 계산
            colCount = colCount + 1
            If colCount >= maxCols Then
                ' 새로운 행으로 이동
                imgTop = imgTop + shp.Height + 10 ' 이미지 높이 + 간격
                imgLeft = ws.Range("K2").Left ' 왼쪽으로 초기화
                colCount = 0
            Else
                ' 오른쪽으로 이동
                imgLeft = imgLeft + shp.Width + 10 ' 이미지 너비 + 간격
            End If
        End If
        
        ' 다음 파일로 이동
        fileName = Dir
    Loop
End Sub


