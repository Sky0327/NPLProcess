Attribute VB_Name = "a2_courtAuction"
Sub UpdateTableAuction()
    DeleteAllRowsInTable "Tpl_Report_법원경매", "tableAuction"
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("3")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행

    Call CopyUniqueIDsToReport
    Call UpdateCourtAndCaseNumber
    Call UpdateExistingRowsInReport
    Call UpdateAuctionStartDate
    Call UpdateAuctionEndDate
    Call UpdateFirstAuctionConclusion
    Call UpdateAuctionFinalPrice
    Call UpdateAuctionFinalDate
    Call UpdateAuctionFinalResult
    Call UpdateFirstAuctionPrice
    Call UpdateFirstAuctionDate
    Call ProcessAuctionData
    
    Call Hyperlink_sheet("3", "Output_법원경매")
    Call AutoFitColumnsExceptA("3")
End Sub

'Report_법원경매의 '등기부등본고유번호' 채우기
Sub CopyUniqueIDsToReport()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim uniqueDict As Object
    Dim tbl As ListObject
    Dim key As Variant
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Dictionary 객체 생성 (중복 제거를 위해)
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    
    ' 중복 제거된 '등기부등본고유번호' 수집
    For i = 2 To lastRow ' Assuming row 1 contains headers
        If Not uniqueDict.Exists(wsInput.Cells(i, 1).value) Then
            uniqueDict.Add wsInput.Cells(i, 1).value, Nothing
        End If
    Next i
    
  
    For Each key In uniqueDict.Keys
        With tbl.ListRows.Add
            .Range(1, 1).value = key ' "등기부등본고유번호" 열에 고유번호 입력
        End With
    Next key
    
    ' 메모리 해제
    Set uniqueDict = Nothing
End Sub

'Report_법원경매의 관할법원과 사건번호 채우기
Sub UpdateCourtAndCaseNumber()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionNumber As String
    Dim splitValues() As String
    Dim tbl As ListObject
    Dim currentID As String
    Dim reportID As String
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            If currentID = reportID Then
                auctionNumber = wsInput.Cells(i, 4).value ' '경매번호' 열 (D열) 기준
                
                ' 경매번호가 특정 패턴을 포함하는지 확인
                If InStr(auctionNumber, "타경") > 0 And InStr(auctionNumber, "_") > 0 Then
                    ' "_"를 기준으로 문자열 분리
                    splitValues = Split(auctionNumber, "_")
                    
                    ' 값 입력
                    tbl.ListRows(j).Range(1, 2).value = splitValues(0) ' "관할법원" 열에 '_' 앞의 텍스트
                    tbl.ListRows(j).Range(1, 3).value = splitValues(1) ' "사건번호" 열에 '_' 뒤의 텍스트
                    
                    ' 일치하는 값을 찾았으므로 내부 루프 종료
                    Exit For
                End If
            End If
        Next i
    Next j
End Sub


'Report_법원경매의 진행상태 채우기
Sub UpdateExistingRowsInReport()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentID As String
    Dim content As String
    Dim dict As Object
    Dim keysArray() As Variant
    Dim j As Long
    Dim tbl As ListObject
    Dim rowCount As Long
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Dictionary 객체 생성 (고유번호별 상태 추적)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 고유번호별 상태 판별
    For i = 2 To lastRow ' Assuming row 1 contains headers
        currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
        content = wsInput.Cells(i, 8).value ' '내용3' 열 (H열) 기준
        
        ' Dictionary에 현재 ID가 없으면 기본 상태로 추가
        If Not dict.Exists(currentID) Then
            dict.Add currentID, "유찰" ' 기본 상태를 "유찰"로 설정
        End If
        
        ' '조회'가 포함된 경우 우선적으로 "조회 내역 없음" 설정
        If InStr(content, "조회") > 0 Then
            dict(currentID) = "조회 내역 없음"
        ElseIf InStr(content, "낙찰") > 0 And dict(currentID) <> "조회 내역 없음" Then
            ' '낙찰'이 포함된 경우 (단, "조회 내역 없음" 상태가 아닌 경우)
            dict(currentID) = "낙찰"
        End If
    Next i
    
    ' Report 시트의 tableAuction 표의 기존 행을 업데이트
    rowCount = tbl.ListRows.Count
    
    For j = 1 To rowCount
        currentID = tbl.ListRows(j).Range(1, 1).value ' "등기부등본고유번호" 열의 값 읽기
        
        ' Dictionary에서 해당 고유번호의 상태를 찾고 "진행상태" 열에 입력
        If dict.Exists(currentID) Then
            tbl.ListRows(j).Range(1, 4).value = dict(currentID) ' "진행상태" 열에 상태 입력
        Else
            tbl.ListRows(j).Range(1, 4).value = "조회 내역 없음" ' 만약 Dictionary에 없으면 기본 상태 입력
        End If
    Next j
    
    ' 메모리 해제
    Set dict = Nothing
End Sub

'Report_법원경매의 '경매개시일' 채우기
Sub UpdateAuctionStartDate()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionStartDate As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, D열의 값이 '경매개시일'인 경우
            If currentID = reportID And wsInput.Cells(i, 4).value = "경매개시일" Then
                ' E열의 '경매개시일' 값을 가져오기
                auctionStartDate = wsInput.Cells(i, 5).value
                
                ' Report 시트의 7번째 열에 값 입력
                tbl.ListRows(j).Range(1, 7).value = auctionStartDate
                
                ' 일치하는 값을 찾았으므로 내부 루프 종료
                Exit For
            End If
        Next i
    Next j
End Sub

'Report_법원경매의 '배당종기일' 채우기
Sub UpdateAuctionEndDate()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionStartDate As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, D열의 값이 '배당종기일'인 경우
            If currentID = reportID And wsInput.Cells(i, 4).value = "배당종기일" Then
                ' E열의 '경매개시일' 값을 가져오기
                auctionStartDate = wsInput.Cells(i, 5).value
                
                ' Report 시트의 8번째 열에 값 입력
                tbl.ListRows(j).Range(1, 8).value = auctionStartDate
                
                ' 일치하는 값을 찾았으므로 내부 루프 종료
                Exit For
            End If
        Next i
    Next j
End Sub

'Report_법원경매의 '최초경매결과' 채우기
Sub UpdateFirstAuctionConclusion()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim specialContent As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim found As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, D열의 값에 "("가 포함된 경우
            If currentID = reportID And InStr(wsInput.Cells(i, 4).value, "(") > 0 Then
                ' H열의 값을 가져오기
                specialContent = wsInput.Cells(i, 8).value
                
                ' Report 시트의 10번째 열에 값 입력
                tbl.ListRows(j).Range(1, 10).value = specialContent
                
                ' 일치하는 값을 찾았으므로 내부 루프 종료
                found = True
                Exit For
            End If
        Next i
        
        ' 일치하는 행을 찾지 못한 경우 빈 값 입력 (Optional)
        If Not found Then
            tbl.ListRows(j).Range(1, 10).value = "조회 내역 없음" ' 찾지 못한 경우 빈 값 입력
        End If
    Next j
End Sub

'Report_법원경매의 '최초 경매기일' 채우기
Sub UpdateFirstAuctionDate()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim specialContent As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim found As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, D열의 값에 "("가 포함된 경우
            If currentID = reportID And InStr(wsInput.Cells(i, 4).value, "(") > 0 Then
                ' D열의 값을 가져오기
                specialContent = wsInput.Cells(i, 4).value
                
                ' Report 시트의 11번째 열에 값 입력
                tbl.ListRows(j).Range(1, 11).value = specialContent
                
                ' 일치하는 값을 찾았으므로 내부 루프 종료
                found = True
                Exit For
            End If
        Next i
        
        ' 일치하는 행을 찾지 못한 경우 빈 값 입력 (Optional)
        If Not found Then
            tbl.ListRows(j).Range(1, 11).value = "조회 내역 없음" ' 찾지 못한 경우 빈 값 입력
        End If
    Next j
End Sub


'Report_법원경매의 '법사가' 채우기
Sub UpdateFirstAuctionPrice()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim specialContent As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim found As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, D열의 값에 "("가 포함된 경우
            If currentID = reportID And InStr(wsInput.Cells(i, 4).value, "(") > 0 Then
                ' G열의 값을 가져오기
                specialContent = wsInput.Cells(i, 7).value
                
                ' Report 시트의 11번째 열에 값 입력
                tbl.ListRows(j).Range(1, 12).value = specialContent
                
                ' 일치하는 값을 찾았으므로 내부 루프 종료
                found = True
                Exit For
            End If
        Next i
        
        ' 일치하는 행을 찾지 못한 경우 빈 값 입력 (Optional)
        If Not found Then
            tbl.ListRows(j).Range(1, 12).value = "조회 내역 없음" ' 찾지 못한 경우 빈 값 입력
        End If
    Next j
End Sub

'최종입찰결과 채우기
Sub UpdateAuctionFinalResult()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionResult As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim lastMatchRow As Long
    Dim found As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        lastMatchRow = 0
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            ' 등기부등본고유번호가 일치하고, E열의 값에 '매각결과'가 포함된 경우
            If currentID = reportID And InStr(wsInput.Cells(i, 5).value, "매각기일") > 0 Then
                lastMatchRow = i ' 현재 행 번호를 저장
                found = True
            End If
        Next i
        
        ' 마지막으로 일치한 행의 H열 값을 가져오기
        If found Then
            auctionResult = wsInput.Cells(lastMatchRow, 8).value ' H열 값 가져오기
            If auctionResult = "" Then
                auctionResult = "조회 내역 없음"
            End If
        Else
            auctionResult = "조회 내역 없음"
        End If
        
        ' Report 시트의 13번째 열에 값 입력
        tbl.ListRows(j).Range(1, 13).value = auctionResult
    Next j
End Sub

'최종입찰기일 채우기
Sub UpdateAuctionFinalDate()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionResult As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim lastMatchRow As Long
    Dim found As Boolean
    Dim lastMatchedRow As Long
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        lastMatchRow = 0
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            If currentID = reportID And (InStr(wsInput.Cells(i, 8).value, "유찰") > 0 Or InStr(wsInput.Cells(i, 8).value, "낙찰") > 0) Then
                lastMatchRow = i ' 현재 행 번호를 저장
                found = True
            End If
        Next i
        
        ' E열에 '매각결과'가 포함된 경우 그 바로 위의 일자를 가져옴
        If found Then
            auctionResult = wsInput.Cells(lastMatchRow, 4).value ' D열 값 가져오기
            If auctionResult = "" Then
                auctionResult = "조회 내역 없음"
            End If
        Else
            auctionResult = "조회 내역 없음"
        End If
        
        ' Report 시트의 14번째 열에 값 입력
        tbl.ListRows(j).Range(1, 14).value = auctionResult
    Next j
End Sub

Sub UpdateAuctionFinalPrice()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim auctionResult As String
    Dim currentID As String
    Dim reportID As String
    Dim tbl As ListObject
    Dim lastMatchRow As Long
    Dim found As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Output_법원경매")
    Set wsReport = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    
    ' "tableAuction"이라는 표 설정
    Set tbl = wsReport.ListObjects("tableAuction")
    
    ' 마지막 행 찾기 (입력 시트에서)
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Report 시트의 tableAuction 표의 기존 행 순회
    For j = 1 To tbl.ListRows.Count
        ' Report 시트의 '등기부등본고유번호' 열 값 가져오기
        reportID = tbl.ListRows(j).Range(1, 1).value
        lastMatchRow = 0
        found = False
        
        ' 입력 시트에서 '등기부등본고유번호'와 일치하는 값 찾기
        For i = 2 To lastRow ' Assuming row 1 contains headers
            currentID = wsInput.Cells(i, 1).value ' '등기부등본고유번호' 열 (A열) 기준
            
            If currentID = reportID And (InStr(wsInput.Cells(i, 8).value, "유찰") > 0 Or InStr(wsInput.Cells(i, 8).value, "낙찰") > 0) Then
                lastMatchRow = i ' 현재 행 번호를 저장
                found = True
            End If
        Next i
        
        ' 마지막으로 일치한 행의 G열 값을 가져오기
        If found Then
            auctionResult = wsInput.Cells(lastMatchRow, 7).value ' G열 값 가져오기
            If auctionResult = "" Then
                auctionResult = "조회 내역 없음"
            End If
        Else
            auctionResult = "조회 내역 없음"
        End If
        
        ' Report 시트의 15번째 열에 값 입력
        tbl.ListRows(j).Range(1, 15).value = auctionResult
    Next j
End Sub

'경매정보 dataSheet >> tempAuction
Sub ProcessAuctionData()

    Dim wsSource As Worksheet
    Dim wsArea As Worksheet
    Dim wsTemp As Worksheet
    Dim tblAuction As ListObject
    Dim rngAreaAuction As Range
    Dim cell As Range
    Dim uniqueItems As Object
    Dim currentRow As Long
    Dim key As Variant
    Dim caseNumber As String
    Dim 진행상태 As String
    Dim 신청채권자 As String
    Dim 중복경매여부 As String
    Dim 경매개시일 As String
    Dim 배당종기일 As String
    Dim 기록열람수령 As String
    Dim 최초경매결과 As String
    Dim 경매기일 As String
    Dim 법사가 As String
    Dim 다음경매결과 As String
    Dim 경매기일2 As String
    Dim 최저가 As String
    Dim newSheetName As String
    Dim caseNumberCol As Integer
    Dim 진행상태Col As Integer
    Dim 신청채권자Col As Integer
    Dim 중복경매여부Col As Integer
    Dim 경매개시일Col As Integer
    Dim 배당종기일Col As Integer
    Dim 기록열람수령Col As Integer
    Dim 최초경매결과Col As Integer
    Dim 경매기일Col As Integer
    Dim 법사가Col As Integer
    Dim 다음경매결과Col As Integer
    Dim 경매기일2Col As Integer
    Dim 최저가Col As Integer
    Dim wsExists As Boolean
    Dim tempArray As Variant

    newSheetName = "3"

    ' tableAuction 표가 있는 시트와 범위 설정
    Set wsSource = ThisWorkbook.Sheets("Tpl_Report_법원경매")
    Set tblAuction = wsSource.ListObjects("tableAuction")

    ' Tpl_report_area 시트 설정 (시트가 존재하는지 확인)
    On Error Resume Next
    Set wsArea = ThisWorkbook.Sheets("Tpl_report_area")
    On Error GoTo 0
    If wsArea Is Nothing Then
        MsgBox "Tpl_report_area 시트를 찾을 수 없습니다."
        Exit Sub
    End If

    ' areaAuction 범위 설정 (범위가 존재하는지 확인)
    On Error Resume Next
    Set rngAreaAuction = wsArea.Range("areaAuction")
    On Error GoTo 0
    If rngAreaAuction Is Nothing Then
        MsgBox "'areaAuction' 범위를 찾을 수 없습니다."
        Exit Sub
    End If

    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = newSheetName

    ' Dictionary 객체 사용 (Collection 대신 Dictionary로 중복키 관리)
    Set uniqueItems = CreateObject("Scripting.Dictionary")

    ' 헤더에서 열 위치를 찾음
    caseNumberCol = tblAuction.ListColumns("사건번호").Index
    진행상태Col = tblAuction.ListColumns("진행상태").Index
    신청채권자Col = tblAuction.ListColumns("신청채권자").Index
    중복경매여부Col = tblAuction.ListColumns("중복경매여부").Index
    경매개시일Col = tblAuction.ListColumns("경매개시일").Index
    배당종기일Col = tblAuction.ListColumns("배당종기일").Index
    기록열람수령Col = tblAuction.ListColumns("기록열람수령").Index
    최초경매결과Col = tblAuction.ListColumns("최초경매결과").Index
    경매기일Col = tblAuction.ListColumns("경매기일").Index
    법사가Col = tblAuction.ListColumns("법사가").Index
    다음경매결과Col = tblAuction.ListColumns("다음경매결과").Index
    경매기일2Col = tblAuction.ListColumns("경매기일2").Index
    최저가Col = tblAuction.ListColumns("최저가").Index

    ' Dictionary 객체 사용 (Collection 대신 Dictionary로 중복키 관리)
    Set uniqueItems = CreateObject("Scripting.Dictionary")
    
    ' 사건번호별로 고유한 사건번호를 관리하고, 여러 값을 저장
    For Each cell In tblAuction.ListColumns(caseNumberCol).DataBodyRange
        caseNumber = cell.value ' 사건번호 열 참조
        진행상태 = cell.offset(0, 진행상태Col - caseNumberCol).value
        신청채권자 = cell.offset(0, 신청채권자Col - caseNumberCol).value
        중복경매여부 = cell.offset(0, 중복경매여부Col - caseNumberCol).value
        경매개시일 = cell.offset(0, 경매개시일Col - caseNumberCol).value
        배당종기일 = cell.offset(0, 배당종기일Col - caseNumberCol).value
        기록열람수령 = cell.offset(0, 기록열람수령Col - caseNumberCol).value
        최초경매결과 = cell.offset(0, 최초경매결과Col - caseNumberCol).value
        경매기일 = cell.offset(0, 경매기일Col - caseNumberCol).value
        법사가 = cell.offset(0, 법사가Col - caseNumberCol).value
        다음경매결과 = cell.offset(0, 다음경매결과Col - caseNumberCol).value
        경매기일2 = cell.offset(0, 경매기일2Col - caseNumberCol).value
        최저가 = cell.offset(0, 최저가Col - caseNumberCol).value
    
        ' 사건번호에 대한 데이터를 Dictionary에 추가
        If Not uniqueItems.Exists(caseNumber) Then
            uniqueItems.Add caseNumber, Array(진행상태, 신청채권자, 중복경매여부, 경매개시일, 배당종기일, 기록열람수령, 최초경매결과, 경매기일, 법사가, 다음경매결과, 경매기일2, 최저가)
        End If
    Next cell

    ' 고유한 사건번호에 대해 areaAuction을 tempAuction 시트에 복사
    currentRow = 2  ' tempAuction 시트의 B2 셀부터 시작
    For Each key In uniqueItems
        ' areaAuction 범위를 복사하여 tempAuction 시트의 해당 위치에 붙여넣기
        rngAreaAuction.Copy wsTemp.Cells(currentRow, 2)

        ' 사건번호 열의 값을 offset(0, 1)에 입력
        wsTemp.Cells(currentRow, 2).offset(0, 1).value = key
        
        ' 진행상태 값을 offset(0, 4)에 입력
        wsTemp.Cells(currentRow, 2).offset(0, 4).value = uniqueItems(key)(0)
        
        ' 신청채권자 값을 offset(1, 3)에 입력
        wsTemp.Cells(currentRow, 2).offset(1, 3).value = uniqueItems(key)(1)
        
        ' 중복경매여부 값을 offset(1, 5)에 입력
        wsTemp.Cells(currentRow, 2).offset(1, 5).value = uniqueItems(key)(2)
        
        ' 경매개시일 값을 offset(3, 1)에 입력
        wsTemp.Cells(currentRow, 2).offset(3, 1).value = uniqueItems(key)(3)
        
        ' 배당종기일 값을 offset(2, 3)에 입력
        wsTemp.Cells(currentRow, 2).offset(2, 3).value = uniqueItems(key)(4)
        
        ' 기록열람수령 값을 offset(2, 5)에 입력
        wsTemp.Cells(currentRow, 2).offset(2, 5).value = uniqueItems(key)(5)
        
        ' 최초경매결과 값을 offset(3, 1)에 입력
        wsTemp.Cells(currentRow, 2).offset(3, 1).value = uniqueItems(key)(6)
        
        ' 경매기일 값을 offset(3, 3)에 입력
        wsTemp.Cells(currentRow, 2).offset(3, 3).value = uniqueItems(key)(7)
        
        ' 법사가 값을 offset(3, 5)에 입력
        wsTemp.Cells(currentRow, 2).offset(3, 5).value = uniqueItems(key)(8)
        
        ' 다음경매결과 값을 offset(4, 1)에 입력
        wsTemp.Cells(currentRow, 2).offset(4, 1).value = uniqueItems(key)(9)
        
        ' 경매기일2 값을 offset(4, 3)에 입력
        wsTemp.Cells(currentRow, 2).offset(4, 3).value = uniqueItems(key)(10)
        
        ' 최저가 값을 offset(4, 5)에 입력
        wsTemp.Cells(currentRow, 2).offset(4, 5).value = uniqueItems(key)(11)

        ' 다음 표를 위해 한 줄 띄우고 아래로 이동
        currentRow = currentRow + rngAreaAuction.Rows.Count + 1
    Next key

    ' 클립보드를 비웁니다.
    Application.CutCopyMode = False

End Sub



