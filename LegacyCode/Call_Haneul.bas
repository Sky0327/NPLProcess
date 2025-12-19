Attribute VB_Name = "Call_Haneul"
Option Explicit


'KB시세 input sheet 생성
Sub makeInputSheetKB()
    Call makeInputSheet("Tpl_Input_KB시세", "Input_KB시세")
    
End Sub

'KB시세 input sheet index값 채우기
Sub copyIndexKB()
    Call CopyInputIndex("Input_KB시세")
    
End Sub
'KB시세 Input sheet 추가정보 채우기
Sub copyKB()
    Call CopyInputKB

End Sub

'불필요한 행 제거
Sub delKB()
    DeleteRowsIfEmptyInKeyColumn "Input_KB시세", "KB시세"
End Sub

'꾸미기
Sub formatKB()
    Call FormatTableBySheet("Input_KB시세")
    
End Sub

'KB시세 Input sheet 전체 작성
Sub writeInputKB()
    Dim userChoice As VbMsgBoxResult
    Dim wb As Workbook
    Dim sheetName As String
    Dim targetSheetName As String
    
    ' 작업 시작 알림 메시지
    'MsgBox "작업을 시작합니다."
    
    ' 현재 워크북 참조
    Set wb = ThisWorkbook
    sheetName = "Input_KB시세"
    targetSheetName = "Output_공시지가(전체)"
    
    ' 1. "Input"시트와 "Output_공시지가(전체)" 시트가 존재하는지 확인
    If Not CheckSheetExists(targetSheetName, wb) Then
        ' 시트가 없으면 메시지 출력 후 종료
        MsgBox "공시지가 조회를 먼저 해주세요."
        Call GoEnd
    End If

    Call CheckRequiredSheet("Input")

    ' 2. Input_KB시세 시트가 존재하는지 확인
    If CheckSheetExists(sheetName, wb) Then
        ' 사용자에게 삭제 여부를 묻는 메시지 박스
        userChoice = MsgBox("'" & sheetName & "' 시트가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")

        If userChoice = vbYes Then
            ' 기존 시트 삭제
            Application.DisplayAlerts = False ' 경고창 표시 비활성화
            wb.Sheets(sheetName).Delete
            Application.DisplayAlerts = True ' 경고창 표시 다시 활성화
        ElseIf userChoice = vbNo Or userChoice = vbCancel Then
            ' 함수 실행 취소
            'MsgBox "함수 실행이 취소되었습니다."
            Call GoEnd
        End If
    End If
    
    ' 3. 시트 생성 및 데이터 입력
    Call makeInputSheetKB
    Call copyIndexKB
    Call copyKB
    Call delKB
    Call formatKB
    Call DeleteColumnByHeader("Input_KB시세", "KB시세")
    ' 4. 작업 완료 메시지
    'MsgBox "모든 작업이 완료되었습니다."
End Sub


'법원경매 input sheet 생성
Sub makeInputSheetCourt()
    Call makeInputSheet("Tpl_Input_법원경매", "Input_법원경매")
    
End Sub

'법원경매 input sheet index값 채우기
Sub copyIndexCourt()
    Call CopyInputIndex("Input_법원경매")
    
End Sub

'법원경매 input sheet 추가정보 채우기
Sub copyCourt()
    Call CopyInputCourt
    
End Sub
'불필요한 행 제거
Sub delCourt()
    DeleteRowsIfEmptyInKeyColumn "Input_법원경매", "법원경매"
End Sub

'꾸미기
Sub formatCourt()
    Call FormatTableBySheet("Input_법원경매")
End Sub

'법원경매 Input sheet 전체 작성
Sub writeInputCourt()
    Dim userChoice As VbMsgBoxResult
    Dim wb As Workbook
    Dim sheetName As String
    
    ' 작업 시작 알림 메시지
    'MsgBox "작업을 시작합니다."
    
    ' 현재 워크북 참조
    Set wb = ThisWorkbook
    sheetName = "Input_법원경매"
    
    Call CheckRequiredSheet("Input")
    
    
    ' Input 시트가 존재하는지 확인
    If CheckSheetExists(sheetName, wb) Then
        ' 사용자에게 삭제 여부를 묻는 메시지 박스
        userChoice = MsgBox("'" & sheetName & "' 시트가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")
        
        If userChoice = vbYes Then
            ' 기존 시트 삭제
            Application.DisplayAlerts = False ' 경고창 표시 비활성화
            wb.Sheets(sheetName).Delete
            Application.DisplayAlerts = True ' 경고창 표시 다시 활성화
        ElseIf userChoice = vbNo Or userChoice = vbCancel Then
            ' 함수 실행 취소
            'MsgBox "함수 실행이 취소되었습니다."
            Call GoEnd
        End If
    End If
    
    ' 시트 생성 및 데이터 입력
    Call makeInputSheetCourt
    Call copyIndexCourt
    Call copyCourt
    Call delCourt
    Call formatCourt
    Call DeleteColumnByHeader("Input_법원경매", "법원경매")
    ' 작업 완료 메시지
    'MsgBox "모든 작업이 완료되었습니다."
End Sub


'인포통계 input sheet 생성
Sub makeInputSheetInfoAnalysis()
    Call makeInputSheet("Tpl_Input_인포통계", "Input_인포통계")
    
End Sub

'인포통계 input sheet index값 채우기
Sub copyIndexInfoAnalysis()
    Call CopyInputIndex("Input_인포통계")
    
End Sub

'인포통계 input sheet 추가정보 채우기
Sub copyInfoAnalysis()
    Call CopyInputInfoAnalysis
    
End Sub

'불필요한 행 제거
Sub delInfoAnalysis()
    DeleteRowsIfEmptyInKeyColumn "Input_인포통계", "인포케어"
End Sub

'꾸미기
Sub formatInfoAnalysis()
    Call FormatTableBySheet("Input_인포통계")
End Sub

'인포통계 Input sheet 전체 작성
Sub writeInputInfoAnalysis()
    Dim userChoice As VbMsgBoxResult
    Dim wb As Workbook
    Dim sheetName As String
    
    ' 작업 시작 알림 메시지
    'MsgBox "작업을 시작합니다."
    
    ' 현재 워크북 참조
    Set wb = ThisWorkbook
    sheetName = "Input_인포통계"
    
    Call CheckRequiredSheet("Input")
    
    
    ' Input 시트가 존재하는지 확인
    If CheckSheetExists(sheetName, wb) Then
        ' 사용자에게 삭제 여부를 묻는 메시지 박스
        userChoice = MsgBox("'" & sheetName & "' 시트가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")
        
        If userChoice = vbYes Then
            ' 기존 시트 삭제
            Application.DisplayAlerts = False ' 경고창 표시 비활성화
            wb.Sheets(sheetName).Delete
            Application.DisplayAlerts = True ' 경고창 표시 다시 활성화
        ElseIf userChoice = vbNo Or userChoice = vbCancel Then
            ' 함수 실행 취소
            'MsgBox "함수 실행이 취소되었습니다."
            Call GoEnd
        End If
    End If
    
    ' 시트 생성 및 데이터 입력
    Call makeInputSheetInfoAnalysis
    Call copyIndexInfoAnalysis
    Call copyInfoAnalysis
    Call delInfoAnalysis
    Call formatInfoAnalysis
    Call DeleteColumnByHeader("Input_인포통계", "인포케어")

    ' 작업 완료 메시지
    'MsgBox "모든 작업이 완료되었습니다."
End Sub


'인포통합 input sheet 생성
Sub makeInputSheetInfoAll()
    Call makeInputSheet("Tpl_Input_인포통합", "Input_인포통합")
    
End Sub

'인포통합 input sheet index값 채우기
Sub copyIndexInfoAll()
    Call CopyInputIndex("Input_인포통합")
    
End Sub

'인포통합 input sheet 추가정보 채우기
Sub copyInfoAll()
    Call CopyInputInfoAll
    
End Sub

'불필요한 행 제거
Sub delInfoAll()
    DeleteRowsIfEmptyInKeyColumn "Input_인포통합", "인포케어"
End Sub

'꾸미기
Sub formatInfoAll()
    Call FormatTableBySheet("Input_인포통합")
End Sub

'인포통합 Input sheet 전체 작성
Sub writeInputInfoAll()
    Dim userChoice As VbMsgBoxResult
    Dim wb As Workbook
    Dim sheetName As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 작업 시작 알림 메시지
    'MsgBox "작업을 시작합니다."
    
    Call CheckRequiredSheet("Input")
    
    
    ' 현재 워크북 참조
    Set wb = ThisWorkbook
    sheetName = "Input_인포통합"
    
    ' Input 시트가 존재하는지 확인
    If CheckSheetExists(sheetName, wb) Then
        ' 사용자에게 삭제 여부를 묻는 메시지 박스
        userChoice = MsgBox("'" & sheetName & "' 시트가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")

        If userChoice = vbYes Then
            ' 기존 시트 삭제
            Application.DisplayAlerts = False ' 경고창 표시 비활성화
            wb.Sheets(sheetName).Delete
            Application.DisplayAlerts = True ' 경고창 표시 다시 활성화
        ElseIf userChoice = vbNo Or userChoice = vbCancel Then
            ' 함수 실행 취소
            'MsgBox "함수 실행이 취소되었습니다."
            Call GoEnd
            Exit Sub
        End If
    End If
    
    ' 시트 생성 및 데이터 입력
    Call makeInputSheetInfoAll
    Call copyIndexInfoAll
    Call copyInfoAll
    Call delInfoAll
    Call formatInfoAll
    Call DeleteColumnByHeader("Input_인포통합", "인포케어")
    Call AdjustColumnWidth("Input_인포통합", 3, 18.5, 15, 55, 10, 10, 10, 10, 10, 10, 10, 10, 10, 12)
    
    ' E열에 값이 있는 경우 P열에 1을 채우기 위한 작업
    Set ws = wb.Sheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    
    For i = 7 To lastRow
        If ws.Cells(i, "E").value <> "" Then
            ws.Cells(i, "P").value = 1
        End If
    Next i

    ' 작업 완료 메시지
    'MsgBox "모든 작업이 완료되었습니다."
End Sub

'인포사례상세 꾸미기
Sub formatInfoDetail()
    Call FormatTableBySheet("Input_인포사례상세")
End Sub
Sub FormatCurrencyAndPercentage()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngCurrency As Range
    Dim rngPercentage As Range
    Dim rngCenterAlign As Range
    Dim cell As Range

    On Error Resume Next
    ' "Input_인포사례상세" 시트 설정
    Set ws = ThisWorkbook.Sheets("Input_인포사례상세")

    ' B열의 7행부터 마지막 연속 데이터 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' J열의 7행부터 마지막 행까지 숫자 서식 적용 (세 자리마다 콤마)
    Set rngCurrency = ws.Range("J7:J" & lastRow)
    rngCurrency.NumberFormat = "#,##0"  ' 숫자 서식 설정 (세 자리 콤마)

    ' K열의 7행부터 마지막 행까지 백분율 서식 적용
    Set rngPercentage = ws.Range("K7:K" & lastRow)
    
    ' 백분율로 변환하기 위해 각 셀의 값을 100으로 나누고 서식 적용
    For Each cell In rngPercentage
        cell.value = cell.value / 100
    Next cell
    rngPercentage.NumberFormat = "0.00%"  ' 소수점 2자리 백분율 서식 적용

    ' P열의 7행부터 마지막 행까지 가운데 정렬 적용
    Set rngCenterAlign = ws.Range("P7:P" & lastRow)
    rngCenterAlign.HorizontalAlignment = xlCenter  ' 가운데 정렬 설정
End Sub



'인포사례상세 input sheet 생성
Sub makeInputSheetInfoDetail_origin()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim rng As Range
    Dim lastCol As Long
    Dim userChoice As VbMsgBoxResult
    Dim wsExisting As Worksheet
    Dim 사건번호Col As Long
    Dim headerRow As Long
    Dim col As Long
    Dim foundHeader As Boolean
    Dim copyRange As Range
    Dim totalColumns As Long
    Dim lastDataCol As Long
    
    Call CheckRequiredSheet("Input")
    
    
    ' 헤더는 A1부터 시작한다고 가정
    headerRow = 1
    
    ' Output_인포통합 시트가 있는지 확인
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("Output_인포통합")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "인포케어 통합검색을 먼저 진행해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 사건번호 열 찾기 (A열부터 마지막 열까지 헤더 확인)
    foundHeader = False
    For col = 1 To wsSource.Cells(1, wsSource.columns.Count).End(xlToLeft).Column
        If wsSource.Cells(headerRow, col).value = "사건번호" Then
            사건번호Col = col
            foundHeader = True
            Exit For
        End If
    Next col
    
    ' 사건번호 열을 찾지 못한 경우
    If Not foundHeader Then
        MsgBox "'사건번호' 열을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' Input_인포사례상세 시트가 이미 존재하는지 확인
    On Error Resume Next
    Set wsExisting = ThisWorkbook.Sheets("Input_인포사례상세")
    On Error GoTo 0
    
    If Not wsExisting Is Nothing Then
        ' 동일한 시트명이 있을 경우 사용자에게 삭제 여부 확인
        userChoice = MsgBox("시트 'Input_인포사례상세'가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")
        
        If userChoice = vbYes Then
            ' 기존 시트 삭제
            Application.DisplayAlerts = False
            wsExisting.Delete
            Application.DisplayAlerts = True
        Else
            ' 작업 종료
            MsgBox "작업이 취소되었습니다.", vbInformation
            Exit Sub
        End If
    End If
    
    ' 새로운 Input_인포사례상세 시트 생성
    Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTarget.Name = "Input_인포사례상세"
    
    ' 첫번째 헤더에서 데이터가 있는 마지막 열 찾기
    lastDataCol = wsSource.Cells(1, wsSource.columns.Count).End(xlToLeft).Column
    
    ' 1행(헤더)에 데이터가 있는 열만 복사하여 Input_인포사례상세 시트의 A1부터 붙여넣기
    Set copyRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastDataCol))
    wsTarget.Cells(1, 1).Resize(1, copyRange.columns.Count).value = copyRange.value
    
    ' Output_인포통합 시트에서 마지막 행 찾기 (사건번호 열을 기준으로)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 사건번호Col).End(xlUp).row
    
    ' Input_인포사례상세 시트의 B7부터 시작 (헤더가 B6이므로 데이터는 B7부터)
    targetRow = 2
    
    ' 사건번호 열을 기준으로 데이터가 있는 행을 확인하고 복사
    For i = 2 To lastRow ' Assuming data starts at row 2 (헤더 제외)
        If wsSource.Cells(i, 사건번호Col).value <> "" Then
            ' 복사할 범위 설정 (행의 1열부터 마지막 데이터가 있는 열까지 복사)
            Set copyRange = wsSource.Range(wsSource.Cells(i, 1), wsSource.Cells(i, lastDataCol))
            
            ' 복사된 데이터를 Input_인포사례상세의 B7부터 붙여넣기 (두 번째 열부터 붙여넣기)
            wsTarget.Cells(targetRow, 1).Resize(1, copyRange.columns.Count).value = copyRange.value
            targetRow = targetRow + 1
        End If
    Next i
    
    ' 복사된 데이터의 마지막 열 찾기
    lastCol = wsTarget.Cells(6, wsTarget.columns.Count).End(xlToLeft).Column
    
    ' '선택' 열 추가 (B6에 '선택' 추가)
    wsTarget.Cells(1, lastCol + 1).value = "조회여부(""V"")"
    
    'MsgBox "데이터가 성공적으로 복사되고 '선택' 열이 추가되었습니다.", vbInformation

End Sub

'인포사례상세 input sheet 생성
Sub makeInputSheetInfoDetail()
    Call makeInputSheetInfoDetail_origin
    Call AdjustColumnWidth("Input_인포사례상세", 3, 18.5, 15, 40, 8.5, 8.5, 12, 14, 14, 14, 8.5, 11, 10, 20, 8.5, 13.5)
    Call ApplyColorFormatting_grey("Input_인포사례상세", "A", "B", "C")
    Call ApplyColorFormatting_orange("Input_인포사례상세", "O")
    Call ApplyColorFormatting_vividorange("Input_인포사례상세", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    Call ApplyAllBorders("Input_인포사례상세")
    Call Output_거리계산_setting("Input_인포사례상세", "Input_인포케어 사례 상세조회")
    Call FormatCurrencyAndPercentage
    
End Sub
