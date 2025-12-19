Attribute VB_Name = "Module_HJ"
Option Explicit
Sub ListSheetNames()
    Dim ws As Worksheet
    Dim i As Integer
    i = 1
    
    ' 결과를 기록할 시트를 선택
    With ThisWorkbook.Sheets("임시_시트목록")
        .Cells.Clear ' 기존 데이터를 지우고 새로 작성
        
        ' 현재 워크북의 모든 시트 이름을 순서대로 출력
        For Each ws In ThisWorkbook.Sheets
            .Cells(i, 1).value = ws.Name
            i = i + 1
        Next ws
    End With
    
End Sub

Sub get_registry_info_basic_selectfolder()

    Dim FolderPicker As FileDialog
    Dim folderPath As String
    
    ' 폴더 선택 창 띄우기
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FolderPicker
        .Title = "폴더를 선택하세요"
        .AllowMultiSelect = False ' 다중 선택 불가
        If .Show = -1 Then ' 사용자가 폴더를 선택했을 경우
            folderPath = .SelectedItems(1) ' 선택된 폴더 경로 저장
            shtSource.Cells(shtSource.columns(1).Find("등기부등본_input폴더경로").row, 2).value = folderPath
        Else
            folderPath = "" ' 폴더 선택이 취소된 경우
        End If
    End With
    
    ' 현재 엑셀 파일 저장
    'ThisWorkbook.Save

End Sub

Sub set_input_main_1()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    
    ' 데이터를 복사할 시트 확인
    On Error GoTo ErrHandler
    Set wsSource = ThisWorkbook.Sheets("Output_등본목록")
    
    ' 데이터를 붙여넣기할 시트 확인
    Set wsTarget = ThisWorkbook.Sheets("Input")
    
    ' 데이터를 복사할 시트에서 데이터 테이블의 마지막 행을 찾음
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    
    ' 제목(1행)을 제외한 A2:E 마지막 행까지의 데이터만 복사
    If lastRow >= 2 Then ' 최소한 한 줄 이상의 데이터가 있는지 확인
        Set dataRange = wsSource.Range("A2:E" & lastRow)
    Else
        MsgBox "데이터가 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 데이터를 붙여넣기할 시트의 7행 A열부터 붙여넣기
    wsTarget.Range("A7").Resize(dataRange.Rows.Count, dataRange.columns.Count).value = dataRange.value
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation
    Exit Sub

' 오류 처리 구간
ErrHandler:
    MsgBox "Output_등본목록 시트가 없어 등기 목록이 생성되지 않았습니다. Input 세팅 - [등기목록 불러오기]를 실행 후 다시 시도해주세요.", vbExclamation
    Exit Sub
End Sub

Sub set_input_2()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim copyRow As Long
    Dim pasteRow As Long
    Dim cell As Range
    
    ' 데이터를 복사할 시트 확인
    On Error GoTo ErrHandler
    Set wsSource = ThisWorkbook.Sheets("Input")
    
    ' 데이터를 붙여넣기할 시트 확인
    Set wsTarget = ThisWorkbook.Sheets("Input_등본조회")
    
    ' 데이터를 복사할 시트에서 데이터 테이블의 마지막 행을 찾음
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row
    
    ' 붙여넣기할 행 시작을 설정
    pasteRow = 7
    
    ' K열에 'V'가 표시된 행에 한해서 데이터 복사
    For Each cell In wsSource.Range("K7:K" & lastRow)
        If cell.value = "V" Then
            ' K열 값이 'V'인 행의 C:E열 값을 복사하여 붙여넣기할 시트로 복사
            wsTarget.Range("B" & pasteRow & ":D" & pasteRow).value = wsSource.Range("C" & cell.row & ":E" & cell.row).value
            ' 서식도 같이 복사
            wsTarget.Range("B7:D7").Copy
            wsTarget.Range("B" & pasteRow & ":D" & pasteRow).PasteSpecial Paste:=xlPasteFormats
            ' 다음 행으로 이동
            pasteRow = pasteRow + 1
        End If
    Next cell
    
    ' 클립보드 비우기
    Application.CutCopyMode = False
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation
    Exit Sub

' 오류 처리 구간
ErrHandler:
    MsgBox "Input 시트를 찾을 수 없습니다. Input 세팅 - [Input 불러오기]를 먼저 실행해주세요.", vbExclamation
    Exit Sub
End Sub

Sub set_input_3()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    Dim cell As Range
    
    ' 데이터를 복사할 시트 확인
    On Error GoTo ErrHandler
    Set wsSource = ThisWorkbook.Sheets("Input")
    
    ' 데이터를 붙여넣기할 시트 확인
    Set wsTarget = ThisWorkbook.Sheets("Input_공시지가")
    
    ' 데이터를 복사할 시트에서 데이터 테이블의 마지막 행을 찾음
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row
    
    ' 붙여넣기할 행 시작을 설정
    pasteRow = 7
    
    ' L열에 'V'가 표시된 행에 한해서 데이터 복사
    For Each cell In wsSource.Range("L7:L" & lastRow)
        If cell.value = "V" Then
            ' L열 값이 'V'인 행의 C:E열 값을 복사하여 붙여넣기할 시트로 복사
            wsTarget.Range("B" & pasteRow & ":D" & pasteRow).value = wsSource.Range("C" & cell.row & ":E" & cell.row).value
            ' 서식도 같이 복사
            wsTarget.Range("B7:D7").Copy
            wsTarget.Range("B" & pasteRow & ":D" & pasteRow).PasteSpecial Paste:=xlPasteFormats
            ' 다음 행으로 이동
            pasteRow = pasteRow + 1
        End If
    Next cell
    
    ' 클립보드 비우기
    Application.CutCopyMode = False
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation
    Exit Sub

' 오류 처리 구간
ErrHandler:
    MsgBox "Input 시트를 찾을 수 없습니다. Input 세팅 - [Input 불러오기]를 먼저 실행해주세요.", vbExclamation
    Exit Sub
End Sub

Sub set_input_4()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    Dim searchColumn As Range
    Dim colIndex As Long
    Dim targetColumnName, targetColumnName2 As String
    Dim cell As Range
    Dim isValid As Boolean
    
    On Error GoTo ErrHandler
    
    ' 찾을 열 이름 정의
    targetColumnName = "실거래가_조회기간(2)"
    targetColumnName2 = "실거래가_구분"
    
    ' 데이터를 복사할 시트 확인
    Set wsSource = ThisWorkbook.Sheets("Input")
    
    ' 데이터를 붙여넣기할 시트 확인
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가")
    
    ' 데이터를 복사할 시트에서 데이터 테이블의 마지막 행을 찾음
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row
    
    ' 유효성 검증
    isValid = True ' 초기값 설정
    
    For Each cell In wsSource.Range("M7:M" & lastRow)
        If cell.value = "V" Then
            ' Y열이 공란이거나, AA열이 "조회기간 선택필요"인 경우 유효성 검증 실패
            If IsEmpty(wsSource.Cells(cell.row, "Y")) Or wsSource.Cells(cell.row, "AA").value = "조회기간 선택필요" Then
                MsgBox "필요한 정보가 입력되지 않았습니다. Input 시트의 [실거래가_구분], [실거래가_조회기간] 열을 확인해주세요.", vbExclamation
                isValid = False ' 유효성 검증 실패
                Exit For ' 유효성 검증 실패 시 루프 종료
            End If
        End If
    Next cell
    
    ' 유효성 검증 실패 시 코드 종료
    If Not isValid Then Exit Sub
    
    ' 붙여넣기할 행 시작을 설정
    pasteRow = 7
    
    ' 유효성 검증 통과 후 데이터 복사 루프 실행
    For Each cell In wsSource.Range("M7:M" & lastRow)
        If cell.value = "V" Then
            ' 복붙1: M열 값이 'V'인 행의 C:F열 데이터를 붙여넣기할 시트의 B열부터 복사
            wsTarget.Range("B" & pasteRow & ":E" & pasteRow).value = wsSource.Range("C" & cell.row & ":F" & cell.row).value
            
            ' 복붙2: 실거래가_조회기간(2) 열의 값을 복사
            Set searchColumn = wsSource.Rows(6).Find(What:=targetColumnName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not searchColumn Is Nothing Then
                colIndex = searchColumn.Column
                wsTarget.Range("G" & pasteRow).value = wsSource.Cells(cell.row, colIndex).value
            Else
                MsgBox "열 '" & targetColumnName & "'을(를) 찾을 수 없습니다.", vbExclamation
                Exit Sub
            End If
            
            ' 복붙3: 실거래가_구분 열의 값을 복사
            Set searchColumn = wsSource.Rows(6).Find(What:=targetColumnName2, LookIn:=xlValues, LookAt:=xlWhole)
            If Not searchColumn Is Nothing Then
                colIndex = searchColumn.Column
                wsTarget.Range("F" & pasteRow).value = wsSource.Cells(cell.row, colIndex).value
            Else
                MsgBox "열 '" & targetColumnName2 & "'을(를) 찾을 수 없습니다.", vbExclamation
                Exit Sub
            End If
            
            ' 다음 행으로 이동
            pasteRow = pasteRow + 1
        End If
    Next cell
    
    ' 붙여넣기할 시트의 7행 서식을 마지막 행까지 적용
    wsTarget.Range("B7:AS7").Copy
    wsTarget.Range("B8:AS" & pasteRow - 1).PasteSpecial Paste:=xlPasteFormats

    ' 클립보드 비우기
    Application.CutCopyMode = False
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation
    Exit Sub

' 오류 처리 구간
ErrHandler:
    MsgBox "Input 시트를 찾을 수 없습니다. Input 세팅 - [Input 불러오기]를 먼저 실행해주세요.", vbExclamation
    Exit Sub
End Sub
Sub MatchAndExtract()

    Dim wsTarget As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowDB As Long
    Dim dbRange As Range
    Dim cell As Range
    Dim matchFound As Boolean
    Dim dbCell As Range
    Dim lastRowTarget As Long
    Dim targetRange As Range

    ' 시트 설정
    Set wsTarget = ThisWorkbook.Sheets("Input") ' 작업 대상 데이터가 있는 시트
    Set wsDB = ThisWorkbook.Sheets("DB_인포케어_지역구분") ' DB 데이터가 있는 시트

    ' DB 데이터의 마지막 행 찾기
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, "F").End(xlUp).row
    Set dbRange = wsDB.Range("F1:F" & lastRowDB)
      
    
    ' E열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "E").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("E7:E" & lastRowTarget)

    ' 작업 대상 데이터 범위 순회
    For Each cell In targetRange
        matchFound = False
        
        ' DB 데이터 순회
        For Each dbCell In dbRange
            ' Check if DB cell value is within the target data cell value
            If InStr(cell.value, dbCell.value) = 1 Then
                wsTarget.Cells(cell.row, 6).value = dbCell.value ' F열에 결과 기록
                matchFound = True
                Exit For
            End If
        Next dbCell
        
        ' 매칭되는 값이 없는 경우 메시지 기록
        If Not matchFound Then
            wsTarget.Cells(cell.row, 6).value = "매칭되는 값이 없습니다" ' F열에 메시지 기록
        End If
    Next cell

    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation

End Sub

Sub MatchAndExtract2()

    Dim wsTarget As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowDB As Long
    Dim dbRange As Range
    Dim cell As Range
    Dim matchFound As Boolean
    Dim dbCell As Range
    Dim lastRowTarget As Long
    Dim targetRange As Range

    ' 시트 설정
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가") ' 작업 대상 데이터가 있는 시트
    Set wsDB = ThisWorkbook.Sheets("DB_밸류맵") ' DB 데이터가 있는 시트

    ' DB 데이터의 마지막 행 찾기
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, "D").End(xlUp).row
    Set dbRange = wsDB.Range("D1:D" & lastRowDB)
    
    ' D열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
    lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("D7:D" & lastRowTarget)

    ' 작업 대상 데이터 범위 순회
    For Each cell In targetRange
        matchFound = False
        
        ' DB 데이터 순회
        For Each dbCell In dbRange
            ' Check if DB cell value is within the target data cell value
            If InStr(cell.value, dbCell.value) = 1 Then
                ' 찾은 값이 존재하는 행의 1번째, 2번째, 3번째 열 값을 H, I, J 열에 반영
                wsTarget.Cells(cell.row, 8).value = wsDB.Cells(dbCell.row, 1).value ' H열에 1번째 열 값 반영
                wsTarget.Cells(cell.row, 9).value = wsDB.Cells(dbCell.row, 2).value ' I열에 2번째 열 값 반영
                wsTarget.Cells(cell.row, 10).value = wsDB.Cells(dbCell.row, 3).value ' J열에 3번째 열 값 반영
                wsTarget.Cells(cell.row, 15).value = wsDB.Cells(dbCell.row, 5).value 'O열에 5번째 열 값 반영
                matchFound = True
                Exit For
            End If
        Next dbCell
        
        ' 매칭되는 값이 없는 경우 메시지 기록
        If Not matchFound Then
            wsTarget.Cells(cell.row, 8).value = "매칭되는 값이 없습니다" ' H열에 메시지 기록
        End If
    Next cell

    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation

End Sub

Sub MatchAndExtract3()

    Dim wsTarget As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowDB As Long
    Dim dbRange As Range
    Dim cell As Range
    Dim matchFound As Boolean
    Dim dbCell As Range
    Dim lastRowTarget As Long
    Dim targetRange As Range

    ' 시트 설정
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가") ' 작업 대상 데이터가 있는 시트
    Set wsDB = ThisWorkbook.Sheets("DB_실거래가_구분") ' DB 데이터가 있는 시트

    ' DB 데이터의 마지막 행 찾기
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).row
    Set dbRange = wsDB.Range("A1:A" & lastRowDB)
    
    ' F열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "F").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("F7:F" & lastRowTarget)

    ' 작업 대상 데이터 범위 순회
    For Each cell In targetRange
        matchFound = False
        
        ' DB 데이터 순회
        For Each dbCell In dbRange
            ' Check if DB cell value is within the target data cell value
            If InStr(cell.value, dbCell.value) = 1 Then
                ' 찾은 값이 존재하는 행의 2번째 열 값을 K 열에 반영
                wsTarget.Cells(cell.row, 11).value = wsDB.Cells(dbCell.row, 2).value ' K열에 2번째 열 값 반영
                matchFound = True
                Exit For
            End If
        Next dbCell
        
        ' 매칭되는 값이 없는 경우 메시지 기록
        If Not matchFound Then
            wsTarget.Cells(cell.row, 11).value = "매칭되는 값이 없습니다" ' K열에 메시지 기록
        End If
    Next cell

    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation

End Sub
  

Sub MarkV()

    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim targetRange As Range

    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가") ' 작업 대상 데이터가 있는 시트
    ' H열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "H").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("L7:L" & lastRowTarget) ' 11번째 열(K열) 7행부터 마지막 행까지 선택

    ' 선택된 범위에 "V" 값 입력
    targetRange.value = "V"

End Sub

  
Sub MarkUniqueValuesByGroup()
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim targetRange As Range
    Dim cell As Range
    Dim uniqueDict_MOLIT As Object ' 국토교통부 그룹의 고유값을 저장하는 딕셔너리
    Dim uniqueDict_ValueMap As Object ' 밸류맵 그룹의 고유값을 저장하는 딕셔너리
    Dim cellValue As String
    Dim combinedValue As String
    Dim groupKey As String
    
    ' 현재 시트를 설정
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가") ' 작업 대상 데이터가 있는 시트
    Set uniqueDict_MOLIT = CreateObject("Scripting.Dictionary")
    Set uniqueDict_ValueMap = CreateObject("Scripting.Dictionary")
    
    ' H열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "H").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("H7:H" & lastRowTarget)
    
    ' targetRange의 각 행에 대해 고유값 처리 (국토교통부: E, F열, 밸류맵: H, I, J열 합)
    For Each cell In targetRange
        ' K열의 값에 따라 그룹 분리
        groupKey = wsTarget.Cells(cell.row, 11).value ' K열은 11번째 열
        
        ' 값이 빈 셀이 아닌 경우에만 처리
        If groupKey = "국토교통부" Then
            ' 국토교통부 그룹: E, F열의 값을 더한 것을 고유값으로 사용
            combinedValue = wsTarget.Cells(cell.row, 5).value & wsTarget.Cells(cell.row, 6).value ' H열과 I열을 합친 값
            ' 고유값이면 L열에 "V" 표시
            If Not uniqueDict_MOLIT.Exists(combinedValue) Then
                uniqueDict_MOLIT.Add combinedValue, True
                wsTarget.Cells(cell.row, 12).value = "V" ' L열은 12번째 열
            End If
            
        ElseIf groupKey = "밸류맵" Then
            ' 밸류맵 그룹: H, I, J, F열의 값을 더한 것을 고유값으로 사용
            combinedValue = wsTarget.Cells(cell.row, 8).value & wsTarget.Cells(cell.row, 9).value & wsTarget.Cells(cell.row, 10).value & wsTarget.Cells(cell.row, 16).value ' H, I, J, F열을 합친 값
            ' 고유값이면 L열에 "V" 표시
            If Not uniqueDict_ValueMap.Exists(combinedValue) Then
                uniqueDict_ValueMap.Add combinedValue, True
                wsTarget.Cells(cell.row, 12).value = "V" ' L열은 12번째 열
            End If
        End If
    Next cell
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation


End Sub

Sub LookupAndCopy()

    Dim wsTarget As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowTarget As Long
    Dim lastRowDB As Long
    Dim targetcell As Range
    Dim dbCell As Range
    Dim matchRow As Long
    
    ' 시트 설정
    Set wsTarget = ThisWorkbook.Sheets("Input") ' 작업 대상 데이터가 있는 시트
    Set wsDB = ThisWorkbook.Sheets("DB_인포케어_지역구분") ' DB 데이터가 있는 시트

    ' Input 시트에서 F열의 마지막 행 찾기 (7행부터)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "F").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    ' DB 시트에서 F열의 마지막 행 찾기
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, "F").End(xlUp).row
    
    ' Input 시트의 F열 7행부터 마지막 행까지 반복
    For Each targetcell In wsTarget.Range("F7:F" & lastRowTarget)
        ' DB 시트의 F열에서 해당 값을 찾음
        Set dbCell = wsDB.Range("F1:F" & lastRowDB).Find(What:=targetcell.value, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' 값을 찾았을 경우
        If Not dbCell Is Nothing Then
            matchRow = dbCell.row
            ' A, B, C 열 값을 G, H, I 열에 복사
            wsTarget.Cells(targetcell.row, "G").value = wsDB.Cells(matchRow, "A").value
            wsTarget.Cells(targetcell.row, "H").value = wsDB.Cells(matchRow, "B").value
            wsTarget.Cells(targetcell.row, "I").value = wsDB.Cells(matchRow, "C").value
        Else
            ' 값을 찾지 못한 경우, G, H, I 열에 공백 기록
            wsTarget.Cells(targetcell.row, "G").value = ""
            wsTarget.Cells(targetcell.row, "H").value = ""
            wsTarget.Cells(targetcell.row, "I").value = ""
        End If
    Next targetcell

    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation

End Sub


Sub SetValuesBasedOnKColumn()
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim targetRange As Range
    Dim cell As Range
    Dim groupKey As String
    
    ' 작업 대상 시트를 설정
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가")
    
    ' K열에서 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "K").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("K7:K" & lastRowTarget)
    
    ' K열의 각 셀을 확인하여 값을 M열에 설정
    For Each cell In targetRange
        groupKey = cell.value ' K열의 값을 가져옴
        
        If groupKey = "국토교통부" Then
            wsTarget.Cells(cell.row, 13).value = 1 ' M열은 13번째 열
        ElseIf groupKey = "밸류맵" Then
            wsTarget.Cells(cell.row, 13).value = 3 ' M열은 13번째 열
        End If
    Next cell
    
    ' 성공 메시지 출력
    'MsgBox "성공!", vbInformation

End Sub

Sub DoubleLookupAndReturnValuesWithLoops()
    Dim wsTarget As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowTarget As Long
    Dim lastRowDB As Long
    Dim targetRange As Range
    Dim cell As Range
    Dim lookupValue As String
    Dim intermediateValue As String
    Dim resultA As Variant
    Dim resultB As Variant
    Dim resultC As Variant
    Dim foundFirst As Boolean
    Dim foundSecond As Boolean
    Dim i As Long, j As Long
    Dim loopCount As Long
    Dim intermediateColumnOffset As Long ' J열부터 열을 증가시킬 오프셋 변수
    Dim targetColumnOffset As Long ' P열부터 S열, V열 등으로 오프셋
    Dim maxColumns As Long ' M열 값에 따른 최대 채울 열 수

    ' 시트 설정
    Set wsTarget = ThisWorkbook.Sheets("Input_실거래가") ' 작업 대상 데이터가 있는 시트
    Set wsDB = ThisWorkbook.Sheets("DB_밸류맵") ' DB 데이터가 있는 시트
    
    ' DB 데이터의 마지막 행 찾기
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, "E").End(xlUp).row
    
    ' O열 7행부터 마지막 데이터가 있는 행 찾기 (작업 대상 데이터 범위)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "O").End(xlUp).row
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRowTarget < 7 Then
        lastRowTarget = 7
    End If
    Set targetRange = wsTarget.Range("O7:O" & lastRowTarget)
    
    ' (1) P7셀부터 AS열의 마지막행까지 값 삭제
    wsTarget.Range("P7:AS" & lastRowTarget).ClearContents
    
    ' targetRange의 각 셀에 대해 이중 Lookup 실행
    For Each cell In targetRange
        lookupValue = cell.value
        foundFirst = False
        foundSecond = False
        
        ' M열 값에 따라 채울 열 결정 (레벨이 1이면 업데이트하지 않음)
        maxColumns = wsTarget.Cells(cell.row, 13).value ' M열 값 가져오기 (13은 M열)
        If maxColumns = 1 Then
            ' 레벨이 1이면 아무 데이터도 업데이트하지 않고 넘어감
            GoTo NextCell
        End If
        
        ' 첫 번째 Lookup: O열의 값을 DB_밸류맵 시트의 E열에서 찾기
        For i = 1 To lastRowDB
            If wsDB.Cells(i, 5).value = lookupValue Then ' E열에서 찾기
                foundFirst = True
                
                ' 총 maxColumns에 따라 다른 열에서 데이터 가져오기
                For loopCount = 2 To maxColumns
                    intermediateColumnOffset = (loopCount - 2) ' 두 번째 세트부터 J, K, L로 이동
                    intermediateValue = wsDB.Cells(i, 10 + intermediateColumnOffset).value ' J, K, L열 값 가져오기
                    
                    ' 두 번째 Lookup: intermediateValue를 DB_밸류맵 시트의 E열에서 다시 찾기
                    For j = 1 To lastRowDB
                        If wsDB.Cells(j, 5).value = intermediateValue Then ' E열에서 찾기
                            foundSecond = True
                            
                            ' A, B, C열의 값을 가져와서 Input_실거래가 시트에 채움
                            resultA = wsDB.Cells(j, 1).value ' A열 값
                            resultB = wsDB.Cells(j, 2).value ' B열 값
                            resultC = wsDB.Cells(j, 3).value ' C열 값
                            
                            ' 두 번째 세트부터는 P, Q, R / S, T, U / V, W, X ...
                            targetColumnOffset = (loopCount - 2) * 3 ' P, S, V ...
                            wsTarget.Cells(cell.row, 16 + targetColumnOffset).value = resultA ' P열 이후
                            wsTarget.Cells(cell.row, 17 + targetColumnOffset).value = resultB ' Q열 이후
                            wsTarget.Cells(cell.row, 18 + targetColumnOffset).value = resultC ' R열 이후
                            
                            Exit For
                        End If
                    Next j
                Next loopCount
                Exit For
            End If
        Next i
        
        ' 첫 번째 또는 두 번째 값을 찾지 못한 경우 처리 (옵션, 필요시 추가 가능)
        If Not foundFirst Then
            ' 첫 번째 Lookup에서 값을 찾지 못한 경우 처리 (필요시 추가)
        ElseIf Not foundSecond Then
            ' 두 번째 Lookup에서 값을 찾지 못한 경우 처리 (필요시 추가)
        End If
        
NextCell:
    Next cell
    
    'MsgBox "성공!"
End Sub




Sub GetOutput()

    Dim outputFilePath As String
    Dim wbSource As Workbook
    Dim wsToCopy As Worksheet
    Dim ws As Worksheet

    ' output 불러오기
    outputFilePath = shtSource.Cells(shtSource.columns(1).Find("등기부등본_output파일경로").row, 2).value ' output 파일 경로
    
    If Dir(outputFilePath) <> "" Then
        ' 파일이 존재할 경우
        On Error GoTo ErrorHandler ' 에러 핸들링 추가
        ' 해당 파일을 열고
        Set wbSource = Workbooks.Open(outputFilePath)
        
        ' 파일의 모든 시트를 복사
        For Each ws In wbSource.Sheets
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count) ' 현재 워크북에 시트 복사
        Next ws
        
        ' 파일 닫기
        wbSource.Close False
        MsgBox "모든 시트가 성공적으로 복사되었습니다."

    Else
        ' 파일이 존재하지 않을 경우
        MsgBox "파일이 확인되지 않습니다. Step1을 실행해주세요."
    End If

    Exit Sub

ErrorHandler:
    MsgBox "파일을 여는 중 오류가 발생했습니다."
    If Not wbSource Is Nothing Then
        wbSource.Close False
    End If

End Sub

Sub CopySheetsFromExternalWorkbook2()

    Dim shtSource As Worksheet
    Dim cell As Range
    Dim searchValue As String
    Dim OutputFolderPath As String
    Dim externalWorkbook As Workbook
    Dim ws As Worksheet
    Dim filePath As String

    ' 시트 설정
    Set shtSource = ThisWorkbook.Sheets("Source")
    
    ' 찾을 값 설정 (예: 보고서명)
    searchValue = "보고서명"
    
    ' A열에서 검색 값 찾기
    For Each cell In shtSource.Range("A:A")
        If cell.value = searchValue Then
            ' B열 값 가져오기
            OutputFolderPath = cell.offset(0, 1).value
            Exit For
        End If
    Next cell
    
    ' OutputFolderPath가 비어 있으면 종료
    If OutputFolderPath = "" Then
        MsgBox "OutputFolderPath가 설정되지 않았습니다."
        Exit Sub
    End If
    
    ' 경로에 있는 파일 열기 (파일명 예시로 file.xlsx 사용)
    filePath = OutputFolderPath & "\file.xlsx"
    
    On Error GoTo FileOpenError
    Set externalWorkbook = Workbooks.Open(filePath)

    ' 외부 파일의 모든 시트 복사
    For Each ws In externalWorkbook.Sheets
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Next ws
    
    ' 외부 파일 닫기 (저장하지 않음)
    externalWorkbook.Close SaveChanges:=False

    MsgBox "모든 시트가 복사되었습니다."

    Exit Sub

FileOpenError:
    MsgBox "파일을 열 수 없습니다. 파일 경로를 확인하세요."

End Sub
Sub CopyFormulaToLastRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Input 시트를 변수에 설정
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' C열 기준으로 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' 마지막 행이 7행보다 작은 경우 7행으로 설정
    If lastRow < 7 Then
        lastRow = 7
    End If
    
    ' K7:P7 셀의 값을 복사하여 K8부터 P 마지막 행까지 붙여넣기
    ws.Range("K7:P7").Copy
    ws.Range("K8:P" & lastRow).PasteSpecial Paste:=xlPasteValues
    
    ' X7 셀의 수식을 복사하여 X8부터 X 마지막 행까지 붙여넣기
    ws.Range("X7").Copy
    ws.Range("X8:X" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
    ' AA7 셀의 수식을 복사하여 AA8부터 AA 마지막 행까지 붙여넣기
    ws.Range("AA7").Copy
    ws.Range("AA8:AA" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
    ' 클립보드 비우기
    Application.CutCopyMode = False
    
    'MsgBox "성공!"
End Sub

' 수식을 복사하는 서브루틴
Sub CopyFormula(ws As Worksheet, col As Long, lastRow As Long)
    ws.Cells(7, col).Copy
    ws.Range(ws.Cells(8, col), ws.Cells(lastRow, col)).PasteSpecial Paste:=xlPasteFormulas
End Sub



Sub InitializeFile()

    Dim ws As Worksheet
    Dim shtSource As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim answer As VbMsgBoxResult

    ' 1. 메시지 박스 띄우기
    answer = MsgBox("파일을 초기화합니다. Source에 입력된 ID/PW 키는 삭제되지 않습니다.", vbOKCancel + vbInformation, "초기화")

    ' 취소를 누르면 매크로 중단
    If answer = vbCancel Then Exit Sub

    ' 2. "Tpl_", "Source" 또는 "DB"가 포함된 시트를 제외한 나머지 시트 삭제
    Application.DisplayAlerts = False ' 시트 삭제 경고 비활성화
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "Tpl_") = 0 And InStr(ws.Name, "Source") = 0 And InStr(ws.Name, "DB") = 0 And InStr(ws.Name, "main") = 0 And InStr(ws.Name, "KeyTest") = 0 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True ' 경고 다시 활성화

    ' 3. Source 시트에서 C열이 "V"인 경우 B열 값 삭제
    Set shtSource = ThisWorkbook.Sheets("Source")
    lastRow = shtSource.Cells(shtSource.Rows.Count, "C").End(xlUp).row
    
    For i = 1 To lastRow
        If shtSource.Cells(i, "C").value = "V" Then
            shtSource.Cells(i, "B").ClearContents
        End If
    Next i

    ' 4. 완료 메시지 박스
    MsgBox "초기화가 완료되었습니다.", vbInformation, "완료"

End Sub

'유효성 검증
Function ValidateTable(sheetName As String, ParamArray checkColumns() As Variant) As Boolean

    Dim ws As Worksheet
    Dim startCell As Range
    Dim lastRow As Long
    Dim currentRow As Long
    Dim col As Variant
    Dim headerRow As Long
    Dim headerName As String
    Dim cellValue As String
    Dim colIndex As Long
    Dim columnLetter As String
    
    On Error GoTo SheetNotFound
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    On Error GoTo ErrorHandler
    
    ' 데이터가 시작되는 셀
    Set startCell = ws.Range("B7")
    
    ' 헤더가 위치한 행 번호
    headerRow = startCell.row - 1
    
    ' 마지막 행 찾기 (B열 기준으로 연속된 데이터)
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    
    ' 만약 B7부터 시작해서 데이터가 없으면 메시지 표시 후 검증 실패
    If lastRow < startCell.row Then
        MsgBox "데이터가 존재하지 않습니다.", vbExclamation, "유효성 검증"
        ValidateTable = False
        Exit Function
    End If
    
    ' 각 행을 순회하며 지정된 컬럼의 빈 셀을 검사
    For currentRow = startCell.row To lastRow
        For Each col In checkColumns
            ' 컬럼이 문자(예: "D")인지 확인
            If VarType(col) = vbString Then
                On Error Resume Next
                colIndex = ws.Range(col & "1").Column
                If Err.Number <> 0 Then
                    MsgBox "유효하지 않은 컬럼명입니다: " & col, vbExclamation, "유효성 검증"
                    ValidateTable = False
                    Exit Function
                End If
                On Error GoTo 0
            Else
                MsgBox "컬럼명은 문자(알파벳)이어야 합니다: " & col, vbExclamation, "유효성 검증"
                ValidateTable = False
                Exit Function
            End If
            
            ' 셀 값과 헤더명 가져오기
            cellValue = Trim(ws.Cells(currentRow, colIndex).value)
            headerName = ws.Cells(headerRow, colIndex).value
            
            ' 컬럼 열(letter) 가져오기
            columnLetter = Split(ws.Cells(1, colIndex).address, "$")(1)
            
            ' 빈 셀인지 확인
            If cellValue = "" Then
                MsgBox headerName & "(" & columnLetter & "열) 은(는) 필수컬럼입니다." & vbCrLf & "값을 모두 채워주세요.", _
                       vbExclamation, "유효성 검증"
                ValidateTable = False
                Exit Function
            End If
        Next col
    Next currentRow
    
    ' 모든 검증을 통과하면 True 반환
    ValidateTable = True
    Exit Function

SheetNotFound:
    MsgBox "시트 '" & sheetName & "'를 찾을 수 없습니다. 선행 코드를 실행하거나 시트 이름을 다시 확인해주세요.", vbExclamation, "유효성 검증"
    ValidateTable = False
    Exit Function

ErrorHandler:
    MsgBox "처리 중 문제가 발생했습니다. 오류내용: " & Err.Description, vbCritical, "유효성 검증"
    ValidateTable = False

End Function


Sub SpeedUp()
    With Application
        .ScreenUpdating = False: .EnableEvents = False: .DisplayAlerts = False: .Calculation = xlCalculationManual: .AutoCorrect.AutoFillFormulasInLists = False
    End With
    ActiveSheet.DisplayPageBreaks = False
End Sub

Sub SpeedDown()
    With Application
        .ScreenUpdating = True: .EnableEvents = True: .DisplayAlerts = True: .Calculation = xlCalculationAutomatic: .AutoCorrect.AutoFillFormulasInLists = True
    End With
End Sub

' 컬럼 너비 조정 및 1번 행 가운데 정렬
Sub AdjustColumnWidth(sheetName As String, ParamArray columnWidths() As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 열너비 조정
    For i = LBound(columnWidths) To UBound(columnWidths)
        ws.columns(i + 1).ColumnWidth = columnWidths(i)
    Next i
    
    ' 1번 행 전체 가운데 정렬
    ws.Rows(1).HorizontalAlignment = xlCenter
End Sub
' 컬럼 헤드 회색 적용
Sub ApplyColorFormatting_grey(sheetName As String, ParamArray columns() As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    Dim col As Range
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 컬럼 서식 적용
    For i = LBound(columns) To UBound(columns)
        ' 해당 컬럼의 1번 행 선택
        Set col = ws.Range(columns(i) & "1")
        
        ' 셀 채우기 색상 회색
        col.Interior.Color = RGB(125, 125, 125)
        
        ' 글씨 색상 #FFFFFF
        col.Font.Color = RGB(255, 255, 255)
        
        ' 볼드체 적용
        col.Font.Bold = True
    Next i
End Sub
' 컬럼 헤드 진한주황색 적용
Sub ApplyColorFormatting_vividorange(sheetName As String, ParamArray columns() As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    Dim col As Range
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 컬럼 서식 적용
    For i = LBound(columns) To UBound(columns)
        ' 해당 컬럼의 1번 행 선택
        Set col = ws.Range(columns(i) & "1")
        
        ' 셀 채우기 색상 진한주황색
        col.Interior.Color = RGB(235, 140, 0)
        
        ' 글씨 색상 #FFFFFF
        col.Font.Color = RGB(255, 255, 255)
        
        ' 볼드체 적용
        col.Font.Bold = True
    Next i
End Sub
' 컬럼 헤드 진한다홍색 적용
Sub ApplyColorFormatting_vividpersimon(sheetName As String, ParamArray columns() As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    Dim col As Range
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 컬럼 서식 적용
    For i = LBound(columns) To UBound(columns)
        ' 해당 컬럼의 1번 행 선택
        Set col = ws.Range(columns(i) & "1")
        
        ' 셀 채우기 색상 진한다홍색
        col.Interior.Color = RGB(208, 74, 2)
        
        ' 글씨 색상 #FFFFFF
        col.Font.Color = RGB(255, 255, 255)
        
        ' 볼드체 적용
        col.Font.Bold = True
    Next i
End Sub
' 컬럼 헤드 주황색 적용
Sub ApplyColorFormatting_orange(sheetName As String, ParamArray columns() As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    Dim col As Range
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 컬럼 서식 적용
    For i = LBound(columns) To UBound(columns)
        ' 해당 컬럼의 1번 행 선택
        Set col = ws.Range(columns(i) & "1")
        
        ' 셀 채우기 색상 주황색
        col.Interior.Color = RGB(255, 182, 0)
        
        ' 글씨 색상 #FFFFFF
        col.Font.Color = RGB(0, 0, 0)
        
        ' 볼드체 적용
        col.Font.Bold = True
    Next i
End Sub
' 폰트사이즈 10 적용
Sub ChangeFontSize(sheetName As String)
    Dim ws As Worksheet
    
    ' 시트 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' 시트가 존재하는지 확인
    If ws Is Nothing Then
        MsgBox "시트 '" & sheetName & "'을(를) 찾을 수 없습니다."
        Exit Sub
    End If
    
    ' 시트 전체 셀의 글꼴 크기를 10으로 변경
    ws.Cells.Font.Size = 10
    
End Sub
' 모든방향 테두리 회색 적용
Sub ApplyAllBorders(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim borderColor As Long
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 시트 활성화
    ws.Activate
    
    ' 눈금선 해제
    ActiveWindow.DisplayGridlines = False
    
    ' 연속된 데이터가 존재하는 마지막 행과 열 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column
    
    ' 데이터 범위 설정
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' 테두리 색상 설정 (RGB 값: #BFBFBF)
    borderColor = RGB(191, 191, 191)
    
    ' 모든 방향 테두리 적용
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Color = borderColor
        .Weight = xlThin
    End With
    
End Sub
Sub ApplyAllBorders_FromA6(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim borderColor As Long
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 시트 활성화
    ws.Activate
    
    ' 눈금선 해제
    ActiveWindow.DisplayGridlines = False
    
    ' A열 6행부터 연속된 데이터가 존재하는 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' 6행부터 연속된 데이터가 존재하는 마지막 열 찾기
    lastCol = ws.Cells(6, ws.columns.Count).End(xlToLeft).Column
    
    ' 데이터 범위 설정 (A열 6행부터 마지막 데이터 셀까지)
    Set dataRange = ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol))
    
    ' 테두리 색상 설정 (RGB 값: #BFBFBF)
    borderColor = RGB(191, 191, 191)
    
    ' 모든 방향 테두리 적용
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Color = borderColor
        .Weight = xlThin
    End With
    
End Sub

Sub UpdateOutputWithInputValues()
    Dim wsOutput As Worksheet
    Dim wsInput As Worksheet
    Dim lastRowOutput As Long
    Dim lastRowInput As Long
    Dim cell As Range
    Dim foundCell As Range
    Dim searchValue As String
    
    ' 시트 설정
    Set wsOutput = ThisWorkbook.Sheets("Output_실거래가_밸류맵")
    Set wsInput = ThisWorkbook.Sheets("Input_실거래가")
    
    ' Output 시트의 마지막 행 찾기
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "B").End(xlUp).row
    
    ' Input 시트의 마지막 행 찾기
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "B").End(xlUp).row
    
    ' Output 시트의 B7셀부터 마지막 행까지 순회
    For Each cell In wsOutput.Range("B7:B" & lastRowOutput)
        searchValue = cell.value
        
        ' Input 시트의 B열에서 검색
        Set foundCell = wsInput.Range("B7:B" & lastRowInput).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' 값을 찾았을 때, 해당 행의 D열 값을 Output 시트의 셀 값으로 변경
        If Not foundCell Is Nothing Then
            cell.value = wsInput.Cells(foundCell.row, "D").value
        End If
    Next cell
End Sub
' 거리계산 sheet에서 보고서 포함여부 컬럼 생성
Sub AddReportIncludeColumn(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim headerCell As Range, dataRange As Range
    
    On Error Resume Next
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 연속된 데이터의 마지막 행과 열 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.Count).End(xlToLeft).Column
    
    ' 새로운 열에 "보고서포함_여부" 헤더 추가
    ws.Cells(1, lastCol + 1).value = "보고서포함_여부"
    
    ' 헤더 셀 포맷 설정 (RGB 색상: (255, 182, 0), 볼드체, 글씨 색상 검정)
    Set headerCell = ws.Cells(1, lastCol + 1)
    With headerCell
        .Interior.Color = RGB(255, 182, 0)
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
    
    ' 데이터 범위 설정 (A1부터 연속된 데이터 범위까지)
    Set dataRange = ws.Range(ws.Cells(2, lastCol + 1), ws.Cells(lastRow, lastCol + 1))
    
    ' 데이터 셀 포맷 설정 (RGB 색상: (251, 226, 213))
    dataRange.Interior.Color = RGB(251, 226, 213)
    
End Sub
' 실거래_국토 시트에서 거래일자 열 생성
Sub AddDealDateColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim yearCol As Range, monthCol As Range, dayCol As Range, dateCol As Range
    Dim i As Long
    Dim dealYear As String, dealMonth As String, dealDay As String, dealDate As String
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets("Output_거리_국토")
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' dealYear, dealMonth, dealDay 열 찾기
    Set yearCol = ws.Rows(1).Find("dealYear", LookAt:=xlWhole)
    Set monthCol = ws.Rows(1).Find("dealMonth", LookAt:=xlWhole)
    Set dayCol = ws.Rows(1).Find("dealDay", LookAt:=xlWhole)
    
    If yearCol Is Nothing Or monthCol Is Nothing Or dayCol Is Nothing Then
        MsgBox "dealYear, dealMonth 또는 dealDay 열이 존재하지 않습니다."
        Exit Sub
    End If
    
    ' dealDay 오른쪽에 dealDate 열 삽입
    ws.columns(dayCol.Column + 1).Insert Shift:=xlToRight
    ws.Cells(1, dayCol.Column + 1).value = "dealDate"
    
    ' 데이터 변환 및 새로운 dealDate 값 입력
    For i = 2 To lastRow
        dealYear = ws.Cells(i, yearCol.Column).value
        dealMonth = ws.Cells(i, monthCol.Column).value
        dealDay = ws.Cells(i, dayCol.Column).value
        
        ' dealMonth가 한 자리일 경우 두 자리로 변환
        If Len(dealMonth) = 1 Then
            dealMonth = "0" & dealMonth
        End If
        
        ' dealDay가 한 자리일 경우 두 자리로 변환
        If Len(dealDay) = 1 Then
            dealDay = "0" & dealDay
        End If
        
        ' yyyy-mm-dd 형식으로 합치기
        dealDate = dealYear & "-" & dealMonth & "-" & dealDay
        
        ' dealDate 열에 값 넣기
        ws.Cells(i, dayCol.Column + 1).value = dealDate
    Next i
End Sub
' Input 시트에서 등본조회 및 공시지가 정보를 lookup 할 수 있도록 header 세팅
Sub 담보물정보_headersetting()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' 값 배열
    Dim headers As Variant
    headers = Array("담보_Mapping", "토지_면적", "토지_대지권비율", "건물구조", "건물내역", "대지권_토지_면적", _
                    "대지권_비율", "대지권_대상_면적", "건물_접수일", "지번", "토지_용도지역명_1", "토지_용도지역명_2", _
                    "토지_공시지가_2024", "토지_공시지가_2023", "집합건물_단지명", "집합건물_동", "집합건물_호", _
                    "집합건물_전용면적", "집합건물_공시지가_24", "건물_대지면적_전체", "건물_대지면적_산정", _
                    "건물_건물연면적_전체", "건물_건물연면적_산정", "건물_공시지가_24")

    ' 값 입력 및 서식 설정
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(6, 28 + i).value = headers(i)
        ws.Cells(6, 28 + i).Font.Bold = True
        
        ' 셀 가운데 정렬
        ws.Cells(6, 28 + i).HorizontalAlignment = xlCenter
        
        If i + 28 = 29 Or i + 28 = 31 Or i + 28 = 32 Or i + 28 = 35 Or i + 28 = 36 Or i + 28 = 37 Or i + 28 = 38 Or i + 28 = 40 Or i + 28 = 41 Then
            ' RGB(255, 182, 0) 배경색
            ws.Cells(6, 28 + i).Interior.Color = RGB(255, 182, 0)
        Else
            ' #B5E6A2 배경색
            ws.Cells(6, 28 + i).Interior.Color = RGB(181, 230, 162)
        End If
    Next i
    
    'AB열부터 AY열까지 열 너비 자동 맞춤
    ws.columns("AB:AY").AutoFit
    
End Sub
' 시트 존재 확인 및 삭제여부Check
Function Check_Sheet(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)

    If Not ws Is Nothing Then
        ' 시트가 존재할 경우, 메시지 박스를 띄운다
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("'" & sheetName & "' 시트가 이미 존재합니다." & vbCrLf & "해당 시트를 삭제 후 작업을 계속하시겠습니까?", vbYesNo + vbQuestion, "시트 삭제 확인")

        ' 유저 응답 확인
        If userResponse = vbYes Then
            ' 유저가 예를 누른 경우 시트를 삭제
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Check_Sheet = True ' 시트를 삭제한 경우 True 반환
        Else
            Check_Sheet = False ' 아니오를 선택한 경우 False 반환
        End If
    Else
        ' 시트가 처음부터 존재하지 않는 경우
        Check_Sheet = True ' 시트가 존재하지 않으므로 True 반환하고 계속 진행
    End If
End Function
' 하이퍼링크 삽입(작업할 시트명, 링크되는 시트명)
Sub Hyperlink_sheet(workSheetName As String, linkSheetName As String)
    Dim ws As Worksheet
    Dim linkWs As Worksheet
    Dim linkAddress As String
    
    ' 첫 번째 시트(작업공간)를 변수로 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(workSheetName)
    Set linkWs = ThisWorkbook.Sheets(linkSheetName)
    On Error GoTo 0
    
    ' 시트가 존재하는지 확인
    If ws Is Nothing Then
        MsgBox "작업 시트가 존재하지 않습니다: " & workSheetName, vbExclamation
        Exit Sub
    End If
    
    If linkWs Is Nothing Then
        MsgBox "링크 시트가 존재하지 않습니다: " & linkSheetName, vbExclamation
        Exit Sub
    End If
    
    ' 작업 시트의 최상단에 행 삽입
    ws.Rows(1).Insert Shift:=xlDown
    
    ' 링크 시트 A1 셀의 주소 가져오기
    linkAddress = "'" & linkSheetName & "'!A1"
    
    ' A1 셀에 하이퍼링크 추가
    ws.Hyperlinks.Add Anchor:=ws.Range("A1"), address:="", SubAddress:=linkAddress, TextToDisplay:="From : " & linkSheetName
    
    ' A1 셀을 좌측 정렬로 설정
    ws.Range("A1").HorizontalAlignment = xlLeft
    
End Sub

Sub GoEnd(Optional StrMsg As String)
    Call SpeedDown
    Application.AutomationSecurity = msoAutomationSecurityByUI
    If StrMsg <> vbNullString Then Msg StrMsg
    End
End Sub

Sub CheckRequiredSheet(sheetName As String)
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    ' 주어진 시트 이름을 선택
    Set ws = ThisWorkbook.Sheets(sheetName)
    Exit Sub
    
ErrorHandler:
    MsgBox "시트 '" & sheetName & "'를 찾을 수 없습니다. 앞선 작업을 먼저 진행 후 다시 실행해주세요.", vbExclamation
    Call GoEnd
End Sub


Sub makeReportSheet_0()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("0")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_report_0", "0")
    
End Sub

Sub removeApo()
    Dim ws As Worksheet
    Dim cell As Range
    Dim targetRange As Range
    Dim lastRow As Long
    Dim i As Long
    
    Call SpeedUp
    
    ' 시트 이름 "0"을 참조
    Set ws = ThisWorkbook.Sheets("0")
    
    ' C7부터 I67 범위를 설정
    Set targetRange = ws.Range("C7:I67")
    
    ' 각 셀을 순회하면서 작은 따옴표가 있으면 수식으로 변환
    For Each cell In targetRange
        ' 셀이 텍스트가 아니라면 수식을 확인
        If cell.HasFormula = False Then
            If Left(cell.Formula, 1) = "=" Then
                ' 수식 앞에 작은 따옴표가 있으면 이를 제거
                cell.Formula = cell.Formula
            End If
        End If
    Next cell
    
    ' C열의 마지막 데이터를 가진 행을 찾음 (7행부터 시작)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' 7행부터 마지막 데이터가 있는 행까지 역순으로 탐색하며 값이 0인 행을 삭제
    For i = lastRow To 7 Step -1
        If ws.Cells(i, "C").value = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
    
    Call SpeedDown
    
End Sub
Sub makeReportSheet_1()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("1")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_report_1", "1")
End Sub

Sub ImportSellList()
    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim baseCell As Range
    Dim baseRow As Long, baseCol As Long
    Dim lastRow As Long
    Dim wsReport1 As Worksheet
    Dim copyRange As Range
    Dim chaJuMyeongCell As Range
    Dim rowToCopy As Long
    Dim pasteRow As Long
    Dim currentRow As Long

    ' 1. 파일 선택 팝업
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "매각리스트 파일을 선택해주세요.")
    If filePath = "False" Then Call GoEnd ' 파일 선택 취소 시 종료

    ' 2. "단독" 시트 선택 - 읽기 전용으로 파일 열기
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error Resume Next
    Set ws = wb.Sheets("단독")

    ' 만약 "단독" 시트가 없다면 첫 번째 시트를 선택
    If ws Is Nothing Then
        Set ws = wb.Sheets(1)
        If ws Is Nothing Then
            MsgBox "시트를 불러오지 못했습니다. 매각리스트 파일이 열려있다면 닫고 다시 실행해주세요.", vbInformation
            wb.Close False
            Exit Sub
        End If
    End If
        On Error GoTo 0
    ' 병합 해제
    ws.Cells.UnMerge

    ' 3. B열에서 "금고정보" 찾기
    Set baseCell = ws.columns("B").Find("금고정보", LookIn:=xlValues, LookAt:=xlWhole)
    If baseCell Is Nothing Then
        MsgBox "'금고정보' not found in column B.", vbExclamation
        wb.Close False
        Exit Sub
    End If

    baseRow = baseCell.row
    baseCol = baseCell.Column

    ' 4. B열에서 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' 5. "차주명" 열 찾기 (baseRow + 1에서)
    Set chaJuMyeongCell = ws.Rows(baseRow + 1).Find("차주명", LookIn:=xlValues, LookAt:=xlWhole)
    If chaJuMyeongCell Is Nothing Then
        MsgBox "'차주명' not found.", vbExclamation
        wb.Close False
        Exit Sub
    End If

    ' 6. 복사할 범위 설정 및 "홍길동"이 있는 행 제외하기
    pasteRow = 4 ' 붙여넣기 시작할 행
    Set wsReport1 = ThisWorkbook.Sheets("1")

    For currentRow = baseRow + 3 To lastRow
        If ws.Cells(currentRow, chaJuMyeongCell.Column).value <> "홍길동" Then
            ' 홍길동이 아닌 경우에만 복사
            wsReport1.Cells(pasteRow, 2).Resize(1, 34).value = ws.Cells(currentRow, baseCol).Resize(1, 34).value
            pasteRow = pasteRow + 1
        End If
    Next currentRow

    ' 마무리
    wb.Close False
End Sub

Sub ConvertToXlsx()

    Dim wb As Workbook
    Dim newWb As Workbook
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim originalPath As String
    Dim savePath As String
    Dim reportName As String
    Dim currentTime As String
    Dim nm As Name
    Dim link As Variant
    Dim i As Integer
    Dim links As Variant
    Dim shtSource As Worksheet

    Call SpeedUp
    

    ' 원본 통합문서의 Source 시트 설정
    Set shtSource = ThisWorkbook.Sheets("Source")

    ' 보고서명 및 현재 시간 설정
    reportName = shtSource.Cells(shtSource.columns(1).Find("보고서명").row, 2).value
    currentTime = Format(Now, "YYYYMMDD_HHMMSS")

    ' 저장 경로와 파일 이름 설정
    originalPath = ThisWorkbook.Path
    savePath = originalPath & "\" & "Smart_NPL_" & reportName & "_" & currentTime & ".xlsx"
    
    ' 새로운 통합문서 생성
    Set newWb = Workbooks.Add
    
    ' 현재 통합문서의 모든 시트를 새 통합문서에 복사 (숨겨진 시트 또는 이름에 "Source"가 포함된 시트는 제외)
    For Each ws In ThisWorkbook.Sheets
        ' 시트가 숨겨져 있지 않고, 시트 이름에 "Source"가 포함되지 않은 경우에만 복사
        If ws.Visible = xlSheetVisible And InStr(1, ws.Name, "Source") = 0 And InStr(1, ws.Name, "Tpl_") = 0 And InStr(1, ws.Name, "main") = 0 Then
            ws.Copy After:=newWb.Sheets(newWb.Sheets.Count)
            
            ' 복사된 시트에 대한 작업
            Set wsNew = newWb.Sheets(newWb.Sheets.Count)
            wsNew.Cells.ClearComments ' 시트 내 모든 셀의 주석 제거
            
            ' 시트에 정의된 이름 제거
            For Each nm In wsNew.Parent.Names
                If InStr(1, nm.RefersTo, wsNew.Name & "!") > 0 Then
                    nm.Delete
                End If
            Next nm
            
            ' 테이블을 범위로 전환 (만약 시트에 테이블이 있을 경우)
            On Error Resume Next
            wsNew.ListObjects(1).Unlist ' 테이블을 범위로 전환
            On Error GoTo 0
        End If
    Next ws
    
    ' 외부 연결 끊기
    links = newWb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            newWb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
    
    ' 통합문서 수준에서 정의된 모든 이름 제거
    For Each nm In newWb.Names
        On Error Resume Next ' 오류 발생 시 무시하고 넘어감
        nm.Delete
        On Error GoTo 0 ' 오류 무시 해제
    Next nm
    
    ' 원본 통합문서와 같은 경로에 xlsx 파일로 저장
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook ' 확장자 xlsx
    Application.DisplayAlerts = True
    
    ' 새 통합문서 닫기
    'newWb.Close SaveChanges:=False
    
    Call SpeedDown
    MsgBox "변환이 완료되었습니다. 파일이 저장되었습니다: " & savePath

End Sub

Sub GetToBePDFName()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets("Output_등본목록")
    
    ' F1 셀에 "파일명(ToBe)" 입력
    ws.Range("F1").value = "파일명(ToBe)"
    
    ' A열 기준 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 2행부터 마지막 행까지 F열에 값 반영
    For i = 2 To lastRow
        ws.Cells(i, "F").value = "[" & ws.Cells(i, "D").value & "] " & ws.Cells(i, "E").value & ".pdf"
    Next i

    MsgBox "등기부등본 ToBe 파일명이 추출되었습니다. 변환을 진행해주세요."
    
End Sub

Sub RenamePDFFiles()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim oldFilePath As String
    Dim newFilePath As String
    Dim folderPath As String
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets("Output_등본목록")
    
    ' A열 기준 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 2행부터 마지막 행까지 파일 이름 변경
    For i = 2 To lastRow
        ' A열에서 기존 파일 경로 가져오기
        oldFilePath = ws.Cells(i, "A").value
        
        ' F열에서 변경될 파일 이름 가져오기
        newFilePath = ws.Cells(i, "F").value
        
        ' 폴더 경로 추출 (파일 경로에서 마지막 "\"까지)
        folderPath = Left(oldFilePath, InStrRev(oldFilePath, "\"))
        
        ' 새 파일 경로 생성
        newFilePath = folderPath & newFilePath
        
        On Error GoTo ErrorHandler
        
        ' 파일 이름 변경
        If Dir(oldFilePath) <> "" Then
            Name oldFilePath As newFilePath
        End If
    Next i
    
    MsgBox "등기부등본 파일명 변환이 완료되었습니다. 재진행시 등본목록 불러오기부터 다시 진행해주세요."
    
ErrorHandler:
    MsgBox "접근오류입니다. PDF가 열려 있거나 이미 사용중인 것 같습니다: " & oldFilePath, vbExclamation
    
End Sub

Function CheckForVInColumn(sheetName As String, columnName As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim foundV As Boolean
    Dim columnAddress As String
    
    ' 시트 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' does not exist."
        CheckForVInColumn = False
        Exit Function
    End If
    
    ' 마지막 행 계산
    lastRow = ws.Cells(ws.Rows.Count, columnName).End(xlUp).row
    
    ' 해당 컬럼에서 V 문자열 검색
    foundV = False
    For Each cell In ws.Range(columnName & "1:" & columnName & lastRow)
        If cell.value = "V" Then
            foundV = True
            Exit For
        End If
    Next cell
    
    ' V 문자열이 없다면 시트 및 컬럼으로 이동 후 메세지 박스 출력
    If Not foundV Then
        Application.GoTo ws.Range(columnName & "1"), True ' 시트 및 컬럼으로 화면 이동
        columnAddress = "'" & "보고서포함_여부" & "'열(" & columnName & "열)"
        MsgBox "'" & sheetName & "' 시트의 " & columnAddress & " 에" & vbCrLf & "선택한 항목이 존재하지 않습니다. 항목을 선택 후 진행해주세요."
        CheckForVInColumn = False
    Else
        CheckForVInColumn = True
    End If
End Function

Sub MoveSheetToLast(sheetName As String)
    Dim ws As Worksheet
    
    ' 시트 설정
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' does not exist."
        Exit Sub
    End If
    
    ' 시트를 가장 마지막으로 이동
    ws.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
End Sub

Function CheckCollateralMapping() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim validMapping As Boolean
    
    ' 'Input' 시트를 설정
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' E열 기준으로 7행부터 마지막 데이터가 있는 행을 탐색
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    
    ' AB열에서 7행부터 마지막 행까지 검사
    validMapping = True
    For Each cell In ws.Range("AB7:AB" & lastRow)
        If InStr(cell.value, "담보-") = 0 Then ' "담보-"가 포함되지 않으면
            validMapping = False
            Exit For
        End If
    Next cell
    
    ' 하나라도 "담보-" 문자열이 포함되지 않으면
    If Not validMapping Then
        ' 화면을 AB6으로 이동
        Application.GoTo ws.Range("U1"), True
        
        ' 메세지 출력 후 False 반환
        MsgBox "'Input' 시트의 '담보_Mapping'열(AB열)에" & vbCrLf & "담보가 정상적으로 Mapping 되지 않았습니다." & vbCrLf & "Mapping을 모두 완료한 후에 진행해주세요."
        CheckCollateralMapping = False
    Else
        CheckCollateralMapping = True
    End If
End Function

Sub ExpandLevel_Info()
    Dim wsInput As Worksheet, wsDB As Worksheet
    Dim lastRow As Long, targetRow As Long, newRow(10) As Long
    Dim col_wsInput_시도 As Integer, col_wsInput_군구 As Integer, col_wsInput_동읍 As Integer
    Dim col_wsDB_J(10) As Integer, col_wsDB_bcode As Integer, col_wsInput_Level As Integer
    Dim searchRow As Long, addLevel As Integer
    Dim str시도 As String, str군구 As String, str동읍 As String
    Dim cell As Range, i As Long, j As Long

    On Error Resume Next

    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Input_인포통합")
    Set wsDB = ThisWorkbook.Sheets("DB_인포케어_지역구분")

    '시트 유효성검증
    If wsInput Is Nothing Then
    MsgBox "Input_인포통합 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
      

    ' 현재 선택된 셀의 행 가져오기
    If Selection.Cells.Count > 1 Then
        MsgBox "하나의 셀만 선택해주세요.", vbExclamation
        Exit Sub
    End If
    targetRow = Selection.row

    ' 헤더행이 6행이므로 선택된 행이 6보다 작으면 종료
    If targetRow <= 6 Then
        MsgBox "확장 대상 행이 있는 셀을 클릭한 후 실행해주세요.", vbExclamation
        Exit Sub
    End If

    ' 열 위치 찾기 (헤더행: 6행)
    For Each cell In wsInput.Rows(6).Cells
        Select Case cell.value
            Case "Level": col_wsInput_Level = cell.Column
            Case "시/도": col_wsInput_시도 = cell.Column
            Case "군/구": col_wsInput_군구 = cell.Column
            Case "동/읍": col_wsInput_동읍 = cell.Column
        End Select
    Next cell

    ' 원본행 데이터 가져오기
    str시도 = wsInput.Cells(targetRow, col_wsInput_시도).value
    str군구 = wsInput.Cells(targetRow, col_wsInput_군구).value
    str동읍 = wsInput.Cells(targetRow, col_wsInput_동읍).value

    ' 유효성 검사
    If str시도 = "" Or str군구 = "" Or str동읍 = "" Then
        MsgBox "시/도, 군/구, 동/읍이 모두 기재된 행을 선택 후 다시 버튼을 눌러주세요.", vbExclamation
        Exit Sub
    End If

    If str동읍 = "전체" Then
        MsgBox "동/읍이 '전체'로 선택되어 있는 경우 인근지역 레벨확장 기능을 사용할 수 없습니다. 구체적인 지역을 선택 후 다시 실행해주세요.", vbExclamation
        Exit Sub
    End If

    ' wsDB의 컬럼 찾기
    Dim jIndex As Integer
    jIndex = 1
    For Each cell In wsDB.Rows(1).Cells
        If cell.value = "J" & jIndex Then
            col_wsDB_J(jIndex) = cell.Column
            jIndex = jIndex + 1
            If jIndex > 10 Then Exit For
        End If
        If InStr(1, cell.value, "법정동", vbTextCompare) > 0 Then col_wsDB_bcode = cell.Column
    Next cell

    ' wsDB에서 검색 (A, B, C열이 시도, 군구, 동읍과 일치하는 행 찾기)
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).row
    searchRow = 0
    For i = 2 To lastRow
        If wsDB.Cells(i, 1).value = str시도 And wsDB.Cells(i, 2).value = str군구 And wsDB.Cells(i, 3).value = str동읍 Then
            searchRow = i
            Exit For
        End If
    Next i

    If searchRow = 0 Then
        MsgBox "일치하는 데이터를 DB에서 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 추가할 행 개수 설정 (1 ~ 10)
    On Error Resume Next
    addLevel = CInt(InputBox("추가할 행의 개수를 입력하세요 (1~10)." & vbCrLf & "현재 선택된 행을 기준으로 인근 범위로 확장합니다.", "레벨확장(인근지역 추가)", 2))
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    If addLevel < 1 Or addLevel > 10 Then
        MsgBox "1~10 사이의 값을 입력해주세요.", vbExclamation
        Exit Sub
    End If

    ' J1 ~ J10 값 가져오기
    Dim J_Value(10) As String
    For i = 1 To addLevel
        J_Value(i) = wsDB.Cells(searchRow, col_wsDB_J(i)).value
    Next i

    ' J1 ~ J10 값을 사용하여 wsDB에서 A, B, C열 값을 찾기
    Dim foundRow(10) As Long
    For i = 1 To 10
        foundRow(i) = 0
    Next i
    For i = 2 To lastRow
        For j = 1 To addLevel
            If wsDB.Cells(i, col_wsDB_bcode).value = J_Value(j) Then foundRow(j) = i
        Next j
    Next i

    ' 행 추가 및 데이터 복제
    wsInput.Rows(targetRow + 1).Resize(addLevel).Insert Shift:=xlDown
    For i = 1 To addLevel
        wsInput.Rows(targetRow).Copy
        wsInput.Rows(targetRow + i).PasteSpecial Paste:=xlPasteAll
        wsInput.Cells(targetRow + i, col_wsInput_Level).value = i
        If foundRow(i) > 0 Then
            wsInput.Cells(targetRow + i, col_wsInput_시도).value = wsDB.Cells(foundRow(i), 1).value
            wsInput.Cells(targetRow + i, col_wsInput_군구).value = wsDB.Cells(foundRow(i), 2).value
            wsInput.Cells(targetRow + i, col_wsInput_동읍).value = wsDB.Cells(foundRow(i), 3).value
        End If
    Next i
    Application.CutCopyMode = False

    ' 새로운 행 번호 저장
    For i = 1 To addLevel
        newRow(i) = targetRow + i
    Next i

    ' 찾은 값 적용
    For i = 1 To addLevel
        If foundRow(i) > 0 Then
            wsInput.Cells(newRow(i), col_wsInput_시도).value = wsDB.Cells(foundRow(i), 1).value
            wsInput.Cells(newRow(i), col_wsInput_군구).value = wsDB.Cells(foundRow(i), 2).value
            wsInput.Cells(newRow(i), col_wsInput_동읍).value = wsDB.Cells(foundRow(i), 3).value
        End If
        wsInput.Cells(newRow(i), col_wsInput_Level).value = i + 1
    Next i

    MsgBox "데이터 추가 및 업데이트 완료!", vbInformation
End Sub



