Attribute VB_Name = "Module_Haneul"
' 시트 생성 및 복사 함수
'-------------------------------------------------------------------
' Function: makeInputSheet
' Description:
'   지정된 템플릿 시트를 복사하여 새로운 입력 시트를 생성하는 함수
'
' Parameters:
'   tplSheetName  - (String) 템플릿 시트 이름
'   inputSheetName - (String) 새롭게 생성할 입력 시트 이름
'
' 기능:
'   1. 템플릿 시트가 존재하는지 확인하고, 없으면 메시지 출력 후 종료
'   2. 동일한 이름의 시트가 존재하면 삭제 후 진행
'   3. 템플릿 시트가 숨겨져 있으면 일시적으로 표시 후 복사
'   4. 복사된 시트의 이름을 inputSheetName으로 변경
'   5. 원래 숨겨져 있던 템플릿 시트는 다시 숨김 처리
'   6. 최종적으로 생성된 시트를 표시
'
' 주의사항:
'   - 기존에 동일한 이름의 시트가 있을 경우 자동으로 삭제됨
'   - 템플릿 시트가 숨겨져 있다면 일시적으로 표시 후 다시 숨겨짐
'-------------------------------------------------------------------

Sub makeInputSheet(tplSheetName As String, inputSheetName As String)
    Dim wb As Workbook
    Dim tplSheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim tplSheetWasHidden As Boolean
    
    ' 현재 작업 중인 엑셀 파일 참조
    Set wb = ThisWorkbook
    
    ' 템플릿 시트 참조
    On Error Resume Next
    Set tplSheet = wb.Sheets(tplSheetName)
    On Error GoTo 0
    
    ' 시트가 존재하지 않을 경우 메시지 출력
    If tplSheet Is Nothing Then
        MsgBox "지정된 템플릿 시트를 찾을 수 없습니다: " & tplSheetName
        Exit Sub
    End If
    
    ' 원하는 새 시트 이름
    newSheetName = inputSheetName
    
    ' 동일한 이름의 시트가 이미 존재하는지 확인
    If CheckSheetExists(newSheetName, wb) Then
        ' 기존 시트 삭제
        Application.DisplayAlerts = False ' 경고창 표시 비활성화
        wb.Sheets(newSheetName).Delete
        Application.DisplayAlerts = True ' 경고창 표시 다시 활성화
    End If
    
    ' 템플릿 시트가 숨겨져 있으면 표시 상태로 변경
    tplSheetWasHidden = False
    If tplSheet.Visible <> xlSheetVisible Then
        tplSheetWasHidden = True
        tplSheet.Visible = xlSheetVisible
    End If
    
    ' 템플릿 시트 복사
    tplSheet.Copy After:=tplSheet
    
    ' 템플릿 시트가 원래 숨겨져 있었다면 다시 숨김
    If tplSheetWasHidden Then
        tplSheet.Visible = xlSheetHidden
    End If
    
    ' 복사된 시트 참조 (복사 후 활성화된 시트)
    Set newSheet = wb.ActiveSheet
    
    ' 복사된 시트의 이름 변경
    newSheet.Name = newSheetName
    
    newSheet.Visible = xlSheetVisible
    
    ' 완료 메시지
    'MsgBox "Input 시트 생성 완료 " & ", 시트명: " & newSheetName
End Sub


' 데이터 복사 함수 (Input 시트는 고정, 대상 시트명 선택 가능)
Sub CopyInputIndex(inputSheetName As String)
    
    Dim wsInput As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowInput As Long
    Dim lastRowTarget As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    ' 현재 워크북 참조
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Input 시트는 고정
    Set wsInput = wb.Sheets("Input")
    ' 대상 시트 참조 (파라미터로 받은 대상 시트명)

    On Error Resume Next
    Set wsTarget = wb.Sheets(inputSheetName)
    On Error GoTo 0

    ' 대상 시트가 존재하지 않을 경우 메시지 출력
    If wsTarget Is Nothing Then
        MsgBox "지정된 대상 시트를 찾을 수 없습니다: " & inputSheetName
        Exit Sub
    End If
    
    ' Input 시트에서 마지막 행을 찾음 (6행부터 데이터가 있으므로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Input 시트에서 등기부등본고유번호, 등기부등본구분, 등기부등본주소 열을 찾음
    Dim colID As Long, colType As Long, colAddress As Long
    colID = Application.Match("등기부등본고유번호", wsInput.Rows(6), 0)
    colType = Application.Match("등기부등본구분", wsInput.Rows(6), 0)
    colAddress = Application.Match("등기부등본주소", wsInput.Rows(6), 0)

    If IsError(colID) Or IsError(colType) Or IsError(colAddress) Then
        MsgBox "등기부등본 관련 열을 찾을 수 없습니다."
        Exit Sub
    End If

    ' Input 시트에서 해당 열의 데이터 범위 설정
    Set sourceRange = Union(wsInput.Range(wsInput.Cells(7, colID), wsInput.Cells(lastRowInput, colID)), _
                            wsInput.Range(wsInput.Cells(7, colType), wsInput.Cells(lastRowInput, colType)), _
                            wsInput.Range(wsInput.Cells(7, colAddress), wsInput.Cells(lastRowInput, colAddress)))

    ' 대상 시트에서 데이터를 입력할 첫 번째 빈 행 찾기 (6행부터)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 2).End(xlUp).row + 1
    If lastRowTarget < 6 Then lastRowTarget = 6

    ' 데이터를 B열부터 붙여넣기 (B열이 2번째 열이므로 targetRange의 열 인덱스 2 설정)
    Set targetRange = wsTarget.Cells(lastRowTarget, 2)
    sourceRange.Copy targetRange

    'MsgBox "Index 데이터가 성공적으로 복사되었습니다."
End Sub

'KB시세 추가데이터 복사함수
Sub CopyInputKB()

    Dim wsInputKB As Worksheet
    Dim wsOutput As Worksheet
    Dim wsLookup As Worksheet
    Dim wsInput As Worksheet ' Input 시트
    Dim lastRowInputKB As Long
    Dim lastRowOutput As Long
    Dim lastRowLookup As Long
    Dim lastRowInput As Long
    Dim inputAddress As String
    Dim outputAddress As String
    Dim lookupKey As String
    Dim lookupValue As String
    Dim kbPrice As String ' KB시세 값을 가져오기 위한 변수
    Dim kbPriceCol As Long ' KB시세 열 번호
    Dim i As Long, j As Long, k As Long
    Dim addressName As String
    Dim netArea As String
    Dim foundAddress As Boolean
    Dim foundNetArea As Boolean
    Dim matchValue As String
    Dim matchRow As Range

    On Error Resume Next

    ' 시트 설정
    Set wsInputKB = ThisWorkbook.Sheets("Input_KB시세")
    Set wsOutput = ThisWorkbook.Sheets("Output_공시지가(전체)")
    Set wsLookup = ThisWorkbook.Sheets("Output_등본조회")
    Set wsInput = ThisWorkbook.Sheets("Input") ' "Input" 시트 설정

    ' Input_KB시세 시트에서 마지막 행 찾기 (B열 기준)
    lastRowInputKB = wsInputKB.Cells(wsInputKB.Rows.Count, 2).End(xlUp).row

    ' Output_공시지가(전체) 시트에서 마지막 행 찾기 (주소 열 기준)
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).row

    ' Output_등본조회 시트에서 마지막 행 찾기
    lastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).row ' 등기부등본고유번호 기준

    ' Input 시트에서 마지막 행 찾기 (C열 기준)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "C").End(xlUp).row

    ' KB시세 열의 위치 찾기 (Input 시트의 6행이 헤더 행)
    kbPriceCol = Application.Match("KB시세", wsInput.Rows(6), 0)

    ' KB시세 열을 찾지 못했을 경우 메시지 출력 후 종료
    If IsError(kbPriceCol) Then
        MsgBox "Input 시트에서 KB시세 열을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' Input_KB시세 시트의 D열부터 루프 시작 (7번째 행부터, 헤더 행은 건너뜀)
    For i = 7 To lastRowInputKB
        inputAddress = wsInputKB.Cells(i, 4).value ' 등기부등본주소 값 가져오기

        ' 데이터 초기화
        foundAddress = False
        foundNetArea = False

        ' Output_공시지가(전체) 시트의 주소와 비교
        For j = 2 To lastRowOutput ' 주소는 2번째 행부터 시작한다고 가정
            outputAddress = wsOutput.Cells(j, 1).value ' Output 시트의 주소 열 값 가져오기

            If inputAddress = outputAddress Then
                ' 구분 열에서 'address_name'과 'prvuseAr'의 값을 찾기
                If wsOutput.Cells(j, 2).value = "address_name" Then
                    addressName = wsOutput.Cells(j, 3).value ' 내용 값 가져오기
                    wsInputKB.Cells(i, 5).value = addressName ' 검색할 주소 열에 값 채우기
                    foundAddress = True ' 주소를 찾음
                End If

                If wsOutput.Cells(j, 2).value = "prvuseAr" Then
                    netArea = wsOutput.Cells(j, 3).value ' 내용 값 가져오기
                    wsInputKB.Cells(i, 6).value = netArea ' 전용면적 열에 값 채우기
                    foundNetArea = True ' 전용면적을 찾음
                End If
            End If
        Next j

        ' 전용면적을 찾지 못한 경우, Output_등본조회 시트에서 lookup
        If Not foundNetArea Then
            matchValue = wsInputKB.Cells(i, 2).value ' wsInputKB의 B열 값

            With Worksheets("Output_등본조회")
                Set matchRow = .columns("A").Find(What:=matchValue, LookIn:=xlValues, LookAt:=xlWhole)
                If Not matchRow Is Nothing Then
                    wsInputKB.Cells(i, 6).value = .Cells(matchRow.row, "G").value ' 일치하는 행의 G열 값 입력
                Else
                    wsInputKB.Cells(i, 6).value = "" ' 일치하는 값이 없을 경우 빈 값 입력
                End If
            End With
        End If

        ' Input 시트에서 KB시세 값 찾기 (C열 기준으로 등기부등본고유번호를 비교)
        For j = 7 To lastRowInput ' Input 시트의 7번째 행부터 시작
            If wsInput.Cells(j, 3).value = wsInputKB.Cells(i, 2).value Then ' C열과 B열 비교
                kbPrice = wsInput.Cells(j, kbPriceCol).value ' Input 시트에서 KB시세 값 가져오기
                wsInputKB.Cells(i, 7).value = kbPrice ' Input_KB시세 시트의 KB시세 열에 값 채우기
                Exit For
            End If
        Next j

        ' 주소를 찾지 못한 경우 등기부등본 주소 입력
        If Not foundAddress Then
            wsInputKB.Cells(i, 5).value = wsInputKB.Cells(i, 4).value
        End If

    Next i

    'MsgBox "업데이트가 완료되었습니다."

End Sub







' Input_법원경매 추가데이터 복사 함수
Sub CopyInputCourt()
    
    Dim wsInput As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowInput As Long
    Dim lastRowTarget As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    ' 현재 워크북 참조
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Input 시트는 고정
    Set wsInput = wb.Sheets("Input")
    
    ' 대상 시트 참조 (파라미터로 받은 대상 시트명)
    On Error Resume Next
    Set wsTarget = wb.Sheets("Input_법원경매")
    On Error GoTo 0

    ' 대상 시트가 존재하지 않을 경우 메시지 출력
    If wsTarget Is Nothing Then
        MsgBox "지정된 대상 시트를 찾을 수 없습니다: 'Input_법원경매'"
        Exit Sub
    End If
    
    ' Input 시트에서 마지막 행을 찾음 (6행부터 데이터가 있으므로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Input 시트에서 법원경매 관련 열 찾음
    Dim courtID As Long, courtYear As Long, courtNumber As Long, courtAuction As Long
    courtID = Application.Match("법원정보", wsInput.Rows(6), 0)
    courtYear = Application.Match("경매연도", wsInput.Rows(6), 0)
    courtNumber = Application.Match("경매번호", wsInput.Rows(6), 0)
    courtAuction = Application.Match("법원경매", wsInput.Rows(6), 0) ' 법원경매 열 추가

    ' 열 찾기 오류 처리
    If IsError(courtID) Or IsError(courtYear) Or IsError(courtNumber) Or IsError(courtAuction) Then
        MsgBox "법원경매정보 관련 열을 찾을 수 없습니다."
        Exit Sub
    End If

    ' Input 시트에서 해당 열의 데이터 범위 설정
    Set sourceRange = Union(wsInput.Range(wsInput.Cells(7, courtID), wsInput.Cells(lastRowInput, courtID)), _
                            wsInput.Range(wsInput.Cells(7, courtYear), wsInput.Cells(lastRowInput, courtYear)), _
                            wsInput.Range(wsInput.Cells(7, courtNumber), wsInput.Cells(lastRowInput, courtNumber)), _
                            wsInput.Range(wsInput.Cells(7, courtAuction), wsInput.Cells(lastRowInput, courtAuction))) ' 법원경매 열 추가

    ' 대상 시트에서 데이터를 입력할 첫 번째 빈 행 찾기 (6행부터)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 5).End(xlUp).row + 1
    If lastRowTarget < 6 Then lastRowTarget = 6

    ' 데이터를 B열부터 붙여넣기 (B열이 2번째 열이므로 targetRange의 열 인덱스 2 설정)
    Set targetRange = wsTarget.Cells(lastRowTarget, 5)
    sourceRange.Copy targetRange

    'MsgBox "법원경매정보 조회목적 데이터가 성공적으로 복사되었습니다."
End Sub


' Input_인포통계 추가데이터 복사 함수
Sub CopyInputInfoAnalysis()

    Dim wsInput As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowInput As Long
    Dim lastRowTarget As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    ' 현재 워크북 참조
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Input 시트는 고정
    Set wsInput = wb.Sheets("Input")
    
    ' 대상 시트 참조 (파라미터로 받은 대상 시트명)
    On Error Resume Next
    Set wsTarget = wb.Sheets("Input_인포통계")
    On Error GoTo 0

    ' 대상 시트가 존재하지 않을 경우 메시지 출력
    If wsTarget Is Nothing Then
        MsgBox "지정된 대상 시트를 찾을 수 없습니다: 'Input_인포통계'"
        Exit Sub
    End If
    
    ' Input 시트에서 마지막 행을 찾음 (6행부터 데이터가 있으므로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Input 시트에서 필요한 열의 인덱스를 찾음
    Dim sido As Long, gungu As Long, dong As Long, useBig As Long, useSmall As Long, infoStatistics As Long
    sido = Application.Match("인포케어_1", wsInput.Rows(6), 0)
    gungu = Application.Match("인포케어_2", wsInput.Rows(6), 0)
    dong = Application.Match("인포케어_3", wsInput.Rows(6), 0)
    useBig = Application.Match("용도(대분류)", wsInput.Rows(6), 0)
    useSmall = Application.Match("용도(소분류)", wsInput.Rows(6), 0)
    
    ' "인포통계" 열 찾기, 오류 처리 포함
    On Error Resume Next
    infoStatistics = Application.Match("인포케어", wsInput.Rows(6), 0)
    On Error GoTo 0
    
    ' 열 찾기 오류 처리
    If IsError(sido) Or IsError(gungu) Or IsError(dong) Or IsError(useBig) Or IsError(useSmall) Or infoStatistics = 0 Then
        MsgBox "필요한 열을 찾을 수 없습니다. 인포케어 또는 인포통계 관련 열을 확인하십시오."
        Exit Sub
    End If

    ' Input 시트에서 해당 열의 데이터 범위 설정 (추가된 인포통계 열 포함)
    Set sourceRange = Union(wsInput.Range(wsInput.Cells(7, sido), wsInput.Cells(lastRowInput, sido)), _
                            wsInput.Range(wsInput.Cells(7, gungu), wsInput.Cells(lastRowInput, gungu)), _
                            wsInput.Range(wsInput.Cells(7, dong), wsInput.Cells(lastRowInput, dong)), _
                            wsInput.Range(wsInput.Cells(7, useBig), wsInput.Cells(lastRowInput, useBig)), _
                            wsInput.Range(wsInput.Cells(7, useSmall), wsInput.Cells(lastRowInput, useSmall)), _
                            wsInput.Range(wsInput.Cells(7, infoStatistics), wsInput.Cells(lastRowInput, infoStatistics))) ' 인포통계 열 추가

    ' 대상 시트에서 데이터를 입력할 첫 번째 빈 행 찾기 (6행부터)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 5).End(xlUp).row + 1
    If lastRowTarget < 6 Then lastRowTarget = 6

    ' 데이터를 붙여넣기 (B열부터 시작)
    Set targetRange = wsTarget.Cells(lastRowTarget, 5)
    sourceRange.Copy targetRange

    'MsgBox "인포케어 통계조회 목적 데이터가 성공적으로 복사되었습니다."
End Sub

' Input_인포통합 추가데이터 복사 함수
Sub CopyInputInfoAll()
    
    Dim wsInput As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowInput As Long
    Dim lastRowTarget As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    ' 현재 워크북 참조
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Input 시트는 고정
    Set wsInput = wb.Sheets("Input")
    
    ' 대상 시트 참조 (파라미터로 받은 대상 시트명)
    On Error Resume Next
    Set wsTarget = wb.Sheets("Input_인포통합")
    On Error GoTo 0

    ' 대상 시트가 존재하지 않을 경우 메시지 출력
    If wsTarget Is Nothing Then
        MsgBox "지정된 대상 시트를 찾을 수 없습니다: 'Input_인포통합'"
        Exit Sub
    End If
    
    ' Input 시트에서 마지막 행을 찾음 (6행부터 데이터가 있으므로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    
    ' Input 시트에서 인포케어 및 기타 관련 열 찾기
    Dim sido As Long, gungu As Long, dong As Long, infoCareCol As Long
    sido = Application.Match("인포케어_1", wsInput.Rows(6), 0)
    gungu = Application.Match("인포케어_2", wsInput.Rows(6), 0)
    dong = Application.Match("인포케어_3", wsInput.Rows(6), 0)
    infoCareCol = Application.Match("인포케어", wsInput.Rows(6), 0) ' 인포케어 열 추가

    ' 열 찾기 오류 처리
    If IsError(sido) Or IsError(gungu) Or IsError(dong) Or IsError(infoCareCol) Then
        MsgBox "인포케어 또는 관련 열을 찾을 수 없습니다."
        Exit Sub
    End If

    ' Input 시트에서 해당 열의 데이터 범위 설정 (인포케어 열 추가)
    Set sourceRange = Union(wsInput.Range(wsInput.Cells(7, sido), wsInput.Cells(lastRowInput, sido)), _
                            wsInput.Range(wsInput.Cells(7, gungu), wsInput.Cells(lastRowInput, gungu)), _
                            wsInput.Range(wsInput.Cells(7, dong), wsInput.Cells(lastRowInput, dong)), _
                            wsInput.Range(wsInput.Cells(7, infoCareCol), wsInput.Cells(lastRowInput, infoCareCol))) ' 인포케어 열 추가

    ' 대상 시트에서 데이터를 입력할 첫 번째 빈 행 찾기 (6행부터)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 5).End(xlUp).row + 1
    If lastRowTarget < 6 Then lastRowTarget = 6

    ' 데이터를 B열부터 붙여넣기 (B열이 2번째 열이므로 targetRange의 열 인덱스 2 설정)
    Set targetRange = wsTarget.Cells(lastRowTarget, 5)
    sourceRange.Copy targetRange

    '날짜 및 용도 입력
    Dim wsInfocare As Worksheet
    Dim i As Long
    
    Dim 기준일자Col As Long
    Dim 인포케어조회개시일Col As Long
    Dim 시작년Col As Long
    Dim 시작월Col As Long
    Dim 시작일Col As Long
    Dim 종료년Col As Long
    Dim 종료월Col As Long
    Dim 종료일Col As Long
    Dim 용도대분류Col As Long
    Dim 용도소분류Col As Long
    Dim targetColumn As Long ' 용도(소분류) 데이터를 넣을 대상 열
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsInfocare = ThisWorkbook.Sheets("Input_인포통합")
    
    ' 헤더가 있는 행 (B6 셀부터 시작되는 표이므로 6번째 행에서 헤더 검색)
    Dim headerRowInput As Long
    Dim headerRowInfocare As Long
    headerRowInput = 6 ' Input 시트에서의 헤더 행
    headerRowInfocare = 6 ' Input_인포케어_통합 시트에서의 헤더 행
    
    ' Input 시트에서 필요한 열의 헤더 찾기
    기준일자Col = GetColumnNumber(wsInput, "기준일자", headerRowInput)
    인포케어조회개시일Col = GetColumnNumber(wsInput, "인포케어_조회개시일", headerRowInput)
    용도소분류Col = GetColumnNumber(wsInput, "용도(소분류)", headerRowInput)
    용도대분류Col = GetColumnNumber(wsInput, "용도(대분류)", headerRowInput)
    
    ' Input_인포케어_통합 시트에서 필요한 열의 헤더 찾기
    시작년Col = GetColumnNumber(wsInfocare, "시작년", headerRowInfocare)
    시작월Col = GetColumnNumber(wsInfocare, "시작월", headerRowInfocare)
    시작일Col = GetColumnNumber(wsInfocare, "시작일", headerRowInfocare)
    종료년Col = GetColumnNumber(wsInfocare, "종료년", headerRowInfocare)
    종료월Col = GetColumnNumber(wsInfocare, "종료월", headerRowInfocare)
    종료일Col = GetColumnNumber(wsInfocare, "종료일", headerRowInfocare)
    
    ' Input 시트에서 마지막 행 찾기 (기준일자 열 기준으로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 기준일자Col).End(xlUp).row

    ' "용도(소분류)" 열의 데이터를 대상 열에 복사할 열 번호 설정 (예: 7번째 열)
    targetColumn = 15 ' 여기에 복사하고 싶은 대상 열 번호로 변경

    For i = 7 To lastRowInput ' 데이터는 6번째 행 다음인 7번째 행부터 시작
        
        ' 기준일자 (Input 시트에서 연/월/일을 추출해 Input_인포케어_통합에 넣기)
        If Not IsDate(wsInput.Cells(i, 인포케어조회개시일Col)) Then
            GoTo DateError
        End If
        ' "2022년", "5월", "1일" 형식으로 복사
        wsInfocare.Cells(lastRowTarget, 시작년Col).value = year(wsInput.Cells(i, 인포케어조회개시일Col).value) & "년"
        wsInfocare.Cells(lastRowTarget, 시작월Col).value = Month(wsInput.Cells(i, 인포케어조회개시일Col).value) & "월"
        wsInfocare.Cells(lastRowTarget, 시작일Col).value = Day(wsInput.Cells(i, 인포케어조회개시일Col).value) & "일"
        
        ' 종료일 (Input 시트에서 연/월/일을 추출해 Input_인포케어_통합에 넣기)
        If Not IsDate(wsInput.Cells(i, 기준일자Col)) Then
            GoTo DateError
        End If
        wsInfocare.Cells(lastRowTarget, 종료년Col).value = year(wsInput.Cells(i, 기준일자Col).value) & "년"
        wsInfocare.Cells(lastRowTarget, 종료월Col).value = Month(wsInput.Cells(i, 기준일자Col).value) & "월"
        wsInfocare.Cells(lastRowTarget, 종료일Col).value = Day(wsInput.Cells(i, 기준일자Col).value) & "일"
        
        ' 용도(대분류)에서 값을 그대로 복사 (괄호와 그 내부 텍스트 삭제하지 않음)
        Dim rawText2 As String
        Dim processedText2 As String
        
        rawText2 = wsInput.Cells(i, 용도대분류Col).value
        If LCase(rawText2) <> "용도(대분류)" Then ' 헤더 텍스트인지 확인하고 건너뜀
            processedText2 = Trim(rawText2)
            wsInfocare.Cells(lastRowTarget, targetColumn).value = processedText2 ' 정확한 위치에 복사
        End If
        
        ' 용도(소분류)에서 값을 그대로 복사 (괄호와 그 내부 텍스트 삭제하지 않음)
        Dim rawText As String
        Dim processedText As String
        
        rawText = wsInput.Cells(i, 용도소분류Col).value
        If LCase(rawText) <> "용도(소분류)" Then ' 헤더 텍스트인지 확인하고 건너뜀
            processedText = Trim(rawText)
            wsInfocare.Cells(lastRowTarget, targetColumn + 1).value = processedText ' 정확한 위치에 복사
        End If
        
        ' 다음 빈 행으로 이동
        lastRowTarget = lastRowTarget + 1
    
    Next i

    ' 에러 핸들링 해제
    Exit Sub

DateError:
    MsgBox "Input 시트에서 '기준일자' 및 '인포케어_조회개시일' 값을 확인해주세요. 날짜 형식의 데이터가 입력되어 있어야 합니다.", vbExclamation, "알림"
    Exit Sub

End Sub


'인포사례상세 Input 데이터 입력
Sub CopyDataToInputDetailsSheet()

    Dim wsOutput As Worksheet
    Dim wsInputDetails As Worksheet
    Dim wsInput As Worksheet
    Dim lastRowOutput As Long
    Dim lastRowInputDetails As Long
    Dim lastRowInput As Long
    Dim i As Long
    
    Dim 등기부등본고유번호Col As Long
    Dim 등기부등본주소Col As Long
    Dim 사건번호Col As Long
    Dim 등기부등본주소InputCol As Long
    Dim 사건번호InputCol As Long
    Dim 등기부등본구분Col As Long
    
    Dim 등기부등본고유번호InputCol As Long
    Dim 등기부등본구분InputCol As Long

    ' 시트 설정
    Set wsOutput = ThisWorkbook.Sheets("Output_인포통합")
    Set wsInputDetails = ThisWorkbook.Sheets("Input_인포사례상세")
    Set wsInput = ThisWorkbook.Sheets("Input") ' Input 시트 참조 추가
    
    ' 헤더가 있는 행 설정 (6행에 헤더가 있고 B6부터 표가 시작)
    Dim headerRow As Long
    headerRow = 6
    
    ' Output_인포통합 시트에서 필요한 열의 헤더 찾기
    등기부등본고유번호Col = GetColumnNumber(wsOutput, "등기부등본고유번호", headerRow)
    등기부등본주소Col = GetColumnNumber(wsOutput, "등기부등본주소", headerRow)
    사건번호Col = GetColumnNumber(wsOutput, "사건번호", headerRow)
    
    ' Input_인포사례상세 시트에서 필요한 열의 헤더 찾기
    등기부등본고유번호Col = GetColumnNumber(wsInputDetails, "등기부등본고유번호", headerRow)
    등기부등본주소InputCol = GetColumnNumber(wsInputDetails, "등기부등본주소", headerRow)
    사건번호InputCol = GetColumnNumber(wsInputDetails, "사건번호", headerRow)
    등기부등본구분Col = GetColumnNumber(wsInputDetails, "등기부등본구분", headerRow)
    
    ' Input 시트에서 필요한 열의 헤더 찾기
    등기부등본고유번호InputCol = GetColumnNumber(wsInput, "등기부등본고유번호", headerRow)
    등기부등본구분InputCol = GetColumnNumber(wsInput, "등기부등본구분", headerRow)
    
    ' Output_인포통합 시트에서 마지막 행 찾기 (고유번호 열 기준으로)
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 등기부등본고유번호Col).End(xlUp).row

    ' Input 시트에서 마지막 행 찾기 (고유번호 열 기준으로)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, 등기부등본고유번호InputCol).End(xlUp).row
    
    ' Input_인포사례상세 시트에서 마지막 행 찾기 (등기부등본고유번호 열 기준으로)
    lastRowInputDetails = wsInputDetails.Cells(wsInputDetails.Rows.Count, 등기부등본고유번호Col).End(xlUp).row
    
    ' 두 시트의 데이터 행 수가 동일하다고 가정하고, 반복문 실행
    For i = 6 To lastRowOutput
        
        ' 고유번호 복사
        wsInputDetails.Cells(i, 등기부등본고유번호Col).value = wsOutput.Cells(i, 등기부등본고유번호Col).value
        
        ' 등기부등본주소 복사
        wsInputDetails.Cells(i, 등기부등본주소InputCol).value = wsOutput.Cells(i, 등기부등본주소Col).value
        
        ' 사건번호 복사
        wsInputDetails.Cells(i, 사건번호InputCol).value = wsOutput.Cells(i, 사건번호Col).value

        ' 고유번호를 기준으로 Input 시트에서 등기부등본구분 값을 Lookup하여 복사
        Dim lookupValue As Variant
        lookupValue = Application.VLookup(wsOutput.Cells(i, 등기부등본고유번호Col).value, wsInput.Range(wsInput.Cells(headerRow, 등기부등본고유번호InputCol), wsInput.Cells(lastRowInput, 등기부등본구분InputCol)), 2, False)
        
        ' 만약 lookupValue가 비어있지 않다면, 등기부등본구분 열에 값 복사
        If Not IsError(lookupValue) Then
            wsInputDetails.Cells(i, 등기부등본구분Col).value = lookupValue
        Else
            wsInputDetails.Cells(i, 등기부등본구분Col).value = "N/A" ' 값이 없을 경우 N/A 처리
        End If

    Next i
    
    'MsgBox "데이터가 성공적으로 복사되었습니다."

End Sub

'Output 파일을 가져오는 함수
Sub CopySheetFromTempFolder(ByVal fileName As String, ByVal newSheetName As String)

    ' 변수 선언
    Dim tempFolderPath As String
    Dim wbSource As Workbook
    Dim wbCurrent As Workbook
    Dim wsToCopy As Worksheet

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim folderPath As String
    
    ' "Source" 시트를 변수에 할당
    Set ws = ThisWorkbook.Sheets("Source")
    
    ' 마지막 행을 찾음
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' A열에서 "보고서폴더경로" 값을 찾음
    For i = 1 To lastRow
        If ws.Cells(i, 1).value = "보고서폴더경로" Then
            ' B열의 값에 "\Temp"를 추가한 경로를 생성
            folderPath = ws.Cells(i, 2).value & "\Temp"

    ' 현재 엑셀 파일 경로 가져오기
    tempFolderPath = folderPath
    
    'MsgBox ("산출물이 다음 폴더에 저장됩니다, : " + folderPath)
    
    ' 현재 엑셀 파일 저장
    Set wbCurrent = ThisWorkbook
    
    ' Temp 폴더에서 파일 열기
    On Error GoTo ErrHandler ' 에러 처리
    Set wbSource = Workbooks.Open(tempFolderPath & fileName)
    
    ' Sheet1을 복사
    Set wsToCopy = wbSource.Sheets("Sheet1")
    wsToCopy.Copy After:=wbCurrent.Sheets(wbCurrent.Sheets.Count)
    
    ' 새 시트명을 파라미터 값으로 설정
    ActiveSheet.Name = newSheetName
    
    ' 원본 파일 닫기
    wbSource.Close SaveChanges:=False
    
    'MsgBox "시트가 성공적으로 복사되고 이름이 변경되었습니다!"
    
    Exit Sub

ErrHandler:
    MsgBox "오류: 파일을 열 수 없습니다. 파일명 또는 경로를 확인하세요.", vbExclamation

End Sub

Sub RemoveRowsIfEmptyInInput(targetSheetName As String, headerName As String)
    Dim wsInput As Worksheet
    Dim wsTarget As Worksheet
    Dim inputLastRow As Long, targetLastRow As Long
    Dim headerInputCol As Long, headerTargetCol As Long
    Dim keyInputCol As Long, keyTargetCol As Long
    Dim keyValue As String
    Dim i As Long, matchRow As Variant
    Dim deleteCount As Long
    Dim inputHeaderRow As Long, targetHeaderRow As Long
    Dim totalRows As Long
    
    ' 시트 정의
    Set wsInput = ThisWorkbook.Sheets("Input") ' Input 시트
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName) ' 타겟 시트
    
    ' Input 시트 및 타겟 시트의 헤더 행 위치
    inputHeaderRow = 6
    targetHeaderRow = 6
    
    ' Input 시트 마지막 행 계산 (B열 기준)
    inputLastRow = wsInput.Cells(wsInput.Rows.Count, "B").End(xlUp).row
    
    ' 타겟 시트 마지막 행 계산 (B열 기준)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).row
    
    ' 처리할 행의 총 개수 계산
    totalRows = targetLastRow - targetHeaderRow
    
    ' 총 작업할 행의 개수를 MsgBox로 알림
    'MsgBox "총 " & totalRows & "개의 행에 대해 작업을 수행합니다.", vbInformation
    
    ' Input 시트에서 해당 헤더명이 있는 열 위치 찾기
    On Error Resume Next
    headerInputCol = Application.WorksheetFunction.Match(headerName, wsInput.Rows(inputHeaderRow), 0)
    On Error GoTo 0
    
    If headerInputCol = 0 Then
        MsgBox "Input 시트에서 '" & headerName & "' 헤더를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' Input 시트에서 '등기부등본고유번호' 열 위치 찾기
    On Error Resume Next
    keyInputCol = Application.WorksheetFunction.Match("등기부등본고유번호", wsInput.Rows(inputHeaderRow), 0)
    On Error GoTo 0
    
    If keyInputCol = 0 Then
        MsgBox "Input 시트에서 '등기부등본고유번호' 열을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 타겟 시트에서 '등기부등본고유번호' 열 위치 찾기
    On Error Resume Next
    keyTargetCol = Application.WorksheetFunction.Match("등기부등본고유번호", wsTarget.Rows(targetHeaderRow), 0)
    On Error GoTo 0
    
    If keyTargetCol = 0 Then
        MsgBox "타겟 시트에서 '등기부등본고유번호' 열을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    deleteCount = 0
    
    ' 타겟 시트에서 B열을 기준으로 행을 역순으로 순환
    For i = targetLastRow To targetHeaderRow + 1 Step -1
        keyValue = wsTarget.Cells(i, keyTargetCol).value
        
        If keyValue <> "" Then
            ' Input 시트에서 해당 key값을 가진 행을 찾기
            On Error Resume Next
            matchRow = Application.Match(keyValue, wsInput.Range(wsInput.Cells(inputHeaderRow + 1, keyInputCol), wsInput.Cells(inputLastRow, keyInputCol)), 0)
            On Error GoTo 0
            
            ' key값을 찾았고, 해당 열의 값이 비어있는 경우
            If Not IsError(matchRow) Then
                If wsInput.Cells(inputHeaderRow + matchRow, headerInputCol).value = "" Then
                    ' 타겟 시트에서 해당 행 삭제
                    wsTarget.Rows(i).Delete
                    deleteCount = deleteCount + 1
                End If
            End If
        End If
    Next i
    
    'MsgBox deleteCount & " rows deleted from " & targetSheetName & " sheet.", vbInformation
End Sub
'Input시트 꾸미기
Sub FormatTableBySheet(sheetName As String)

    Dim targetSheet As Worksheet
    Dim rngTable As Range
    Dim rngHeader As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim headerCell As Range
    Dim dataRange As Range
    Dim 조회여부Column As Long
    
    ' 1. 시트명을 파라미터로 받아와서 targetSheet에 할당
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "시트 '" & sheetName & "'을(를) 찾을 수 없습니다."
        Exit Sub
    End If

    ' 2. B6부터 시작하는 표의 범위 인식 (헤더는 6행)
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 2).End(xlUp).row ' B열 기준 마지막 행 찾기
    lastCol = targetSheet.Cells(6, targetSheet.columns.Count).End(xlToLeft).Column ' 6행 기준 마지막 열 찾기
    Set rngTable = targetSheet.Range(targetSheet.Cells(6, 2), targetSheet.Cells(lastRow, lastCol))
    Set rngHeader = targetSheet.Range(targetSheet.Cells(6, 2), targetSheet.Cells(6, lastCol)) ' 6행(헤더)
    
    ' 초기화: 조회여부 열의 위치를 0으로 설정
    조회여부Column = 0
    
    ' 3. 모양 지정
    ' 3-1. 특정 헤더명에 대해 배경 및 글씨체 설정
    For Each cell In rngHeader
        Select Case cell.value
            Case "등기부등본고유번호", "등기부등본구분", "등기부등본주소"
                cell.Interior.Color = RGB(125, 125, 125) ' 짙은 회색
                cell.Font.Color = RGB(255, 255, 255) ' 하얀색 글씨
                cell.Font.Bold = True ' 볼드 처리
            Case "조회여부(""V"")"
                cell.Interior.Color = RGB(255, 182, 0) ' 옅은 주황색
                cell.Font.Color = RGB(0, 0, 0) ' 검정색 글씨
                cell.Font.Bold = True ' 볼드 처리
                조회여부Column = cell.Column ' "조회여부" 열 위치 기록
            Case Else
                cell.Interior.Color = RGB(235, 140, 0) ' 짙은 주황색
                cell.Font.Color = RGB(255, 255, 255) ' 하얀색 글씨
                cell.Font.Bold = True ' 볼드 처리
        End Select
    Next cell
    
    ' 3-2. "조회여부" 열의 데이터 셀에 대해 옅은 노란색 배경 처리
    If 조회여부Column <> 0 Then ' "조회여부" 열이 존재하는 경우
        Set dataRange = targetSheet.Range(targetSheet.Cells(7, 조회여부Column), targetSheet.Cells(lastRow, 조회여부Column))
        dataRange.Interior.Color = RGB(255, 255, 153) ' 옅은 노란색
    End If
    
    ' 3-3. 모든 셀에 대하여 테두리는 옅은 회색
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200) ' 옅은 회색
        .Weight = xlThin
    End With
    
    ' 3-4. 모든 셀의 글꼴을 '맑은 고딕', 글자 크기를 10으로 설정
    With rngTable.Font
        .Name = "맑은 고딕" ' 글꼴 설정
        .Size = 10 ' 글자 크기 설정
    End With
    
    ' 4. 셀눈금 보이지 않게 설정
    targetSheet.Activate
    ActiveWindow.DisplayGridlines = False

    'MsgBox "모양 설정이 완료되었습니다."

End Sub

Sub FormatTableByOutputSheet(sheetName As String)

    Dim targetSheet As Worksheet
    Dim rngTable As Range
    Dim rngHeader As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range
    
    ' 1. 시트명을 파라미터로 받아와서 targetSheet에 할당
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "시트 '" & sheetName & "'을(를) 찾을 수 없습니다."
        Exit Sub
    End If

    ' 2. A1부터 시작하는 표의 범위 인식 (헤더는 1행)
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row ' A열 기준 마지막 행 찾기
    lastCol = targetSheet.Cells(1, targetSheet.columns.Count).End(xlToLeft).Column ' 1행 기준 마지막 열 찾기
    
    Set rngTable = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, lastCol))
    Set rngHeader = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(1, lastCol)) ' 1행(헤더)

    ' 3. 모양 지정
    ' 3-1. 헤더(A1부터 첫 번째 행)의 배경 및 글씨체 설정
    For Each cell In rngHeader
        Select Case cell.value
            Case "등기부등본고유번호", "등기부등본구분", "등기부등본주소"
                cell.Interior.Color = RGB(125, 125, 125) ' 짙은 회색
                cell.Font.Color = RGB(255, 255, 255) ' 하얀색 글씨
                cell.Font.Bold = True ' 볼드 처리
            Case Else
                cell.Interior.Color = RGB(208, 74, 2) ' 짙은 빨간색
                cell.Font.Color = RGB(255, 255, 255) ' 하얀색 글씨
                cell.Font.Bold = True ' 볼드 처리
        End Select
    Next cell
    
    ' 3-2. 모든 셀에 대하여 테두리는 옅은 회색
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200) ' 옅은 회색
        .Weight = xlThin
    End With
    
    ' 3-3. 모든 셀의 글꼴을 '맑은 고딕', 글자 크기를 10으로 설정
    With rngTable.Font
        .Name = "맑은 고딕" ' 글꼴 설정
        .Size = 10 ' 글자 크기 설정
    End With

    ' 4. 셀눈금 보이지 않게 설정
    targetSheet.Activate
    ActiveWindow.DisplayGridlines = False

    'MsgBox "모양 설정이 완료되었습니다."

End Sub


'시트명과 keyColumn명을 받아서 해당 column에 값이 없는 경우 행을 삭제하는 함수.
Sub DeleteRowsIfEmptyInKeyColumn(sheetName As String, keyColumn As String)

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim keyColIndex As Long
    Dim i As Long

    ' 시트 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' 시트를 찾지 못한 경우
    If ws Is Nothing Then
        MsgBox "지정된 시트를 찾을 수 없습니다: " & sheetName, vbExclamation
        Exit Sub
    End If

    ' 헤더가 위치한 행 번호 (6행)
    headerRow = 6

    ' 마지막 행 계산 (B열 기준)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' keyColumn에 해당하는 열 찾기
    On Error Resume Next
    keyColIndex = Application.WorksheetFunction.Match(keyColumn, ws.Rows(headerRow), 0)
    On Error GoTo 0

    ' keyColumn을 찾을 수 없는 경우 메시지 출력 후 종료
    If keyColIndex = 0 Then
        MsgBox "지정된 '" & keyColumn & "' 열을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 마지막 행부터 역순으로 순환하여 빈 셀을 찾아 삭제
    For i = lastRow To headerRow + 1 Step -1
        If ws.Cells(i, keyColIndex).value = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    'MsgBox "빈 셀이 있는 행이 모두 삭제되었습니다.", vbInformation

End Sub

Sub AutoFitColumnsExceptA(sheetName As String)
    Dim ws As Worksheet
    Dim lastCol As Integer
    Dim col As Integer
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 마지막 열 찾기
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column
    
    ' A열 너비를 3으로 설정
    ws.columns(1).ColumnWidth = 3
    
    ' A열을 제외하고 나머지 열의 셀 너비 자동 맞춤
    For col = 2 To lastCol
        ws.columns(col).AutoFit
    Next col
End Sub


Sub AutoFitColumnsBySheetName(sheetName As String)
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    
    ' 시트가 존재하는지 확인
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "해당 시트가 존재하지 않습니다: " & sheetName
        Exit Sub
    End If
    
    ' B열부터 마지막 열까지 확인
    With ws
        lastCol = .Cells(1, .columns.Count).End(xlToLeft).Column
        
        For col = 2 To lastCol ' B열부터 시작 (B열은 2번째 열)
            .columns(col).AutoFit
        Next col
    End With
    
End Sub


Sub HideSheetByName(sheetName As String)
    Dim ws As Worksheet
    
    ' 시트가 존재하는지 확인
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "해당 시트가 존재하지 않습니다: " & sheetName
        Exit Sub
    End If
    
    ' 시트를 숨김(Hidden) 처리
    ws.Visible = xlSheetHidden

End Sub




