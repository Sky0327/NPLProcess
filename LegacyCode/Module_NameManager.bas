Attribute VB_Name = "Module_NameManager"
'설명: 인포케어 지역선택 3단계 드롭다운을 위하여, 데이터범위에 대하여 이름을 부여하는 코드임.
'한번만 실행하면 되나, 추후 오류 등으로 인해 재작업해야 할 수도 있으므로 코드 남겨둠.
'이름관리자에서 List1, List2, List3 으로 시작하는 이름들이 대상임.
'실행하기 전에는 오류방지를 위해 기존 이름은 일괄 삭제하고 실행 권장.

'List1 이름을 지정하는 코드는 없음 (시군구 Group은 1개라서 수기로 하면 됨)

'List2 이름을 지정하는 코드
Sub CreateDynamicNamedRanges2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCity As String
    Dim currentDistrict As String
    Dim startRow As Long
    Dim nameStr As String
    Dim rng As Range
    
    ' 데이터가 있는 시트 설정
    Set ws = ThisWorkbook.Sheets("DB_인포케어_드롭다운1")
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(Rows.Count, 3).End(xlUp).row ' C열
    
    ' 첫 번째 데이터 시작 행
    startRow = 2
    
    ' 루프 실행 (2행부터 마지막 행까지)
    For i = 2 To lastRow + 1
        ' 현재 행의 시도명과 시군구 가져오기
        currentCity = ws.Cells(i, 3).value  ' 시도명 (C열)
        
        ' 다음 행과 비교하여 그룹 변경 확인
        If currentCity <> ws.Cells(i - 1, 3).value Then 'C열
            ' 이름 정의 (첫 번째 행이 아닐 경우)
            If i > 2 Then
                ' 이름 생성 (띄어쓰기 "_" 처리)
                nameStr = "List2_" & ws.Cells(i - 1, 3).value 'C열
                nameStr = Replace(nameStr, " ", "_") ' 공백 제거
                nameStr = Replace(nameStr, "-", "_") ' 특수문자 대체
                
                ' 범위 설정
                Set rng = ws.Range("D" & startRow & ":D" & (i - 1)) 'D열
                
                ' 이름 정의
                ThisWorkbook.Names.Add Name:=nameStr, RefersTo:=rng
            End If
            
            ' 새 그룹 시작
            startRow = i
        End If
    Next i
    
    MsgBox "모든 범위에 이름이 자동으로 부여되었습니다!", vbInformation, "완료"
End Sub

'List3 이름을 지정하는 코드
Sub CreateDynamicNamedRanges3()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCity As String
    Dim currentDistrict As String
    Dim startRow As Long
    Dim nameStr As String
    Dim rng As Range
    
    ' 데이터가 있는 시트 설정
    Set ws = ThisWorkbook.Sheets("DB_인포케어_드롭다운1")
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(Rows.Count, 8).End(xlUp).row ' H열
    
    ' 첫 번째 데이터 시작 행
    startRow = 2
    
    ' 루프 실행 (2행부터 마지막 행까지)
    For i = 2 To lastRow + 1
        ' 현재 행의 시도명과 시군구 가져오기
        currentCity = ws.Cells(i, 6).value  ' 시도명 (F열)
        currentDistrict = ws.Cells(i, 7).value ' 시군구 (G열)
        
        ' 다음 행과 비교하여 그룹 변경 확인
        If currentCity & currentDistrict <> ws.Cells(i - 1, 6).value & ws.Cells(i - 1, 7).value Then 'F열, G열
            ' 이름 정의 (첫 번째 행이 아닐 경우)
            If i > 2 Then
                ' 이름 생성 (띄어쓰기 "_" 처리)
                nameStr = "List3_" & ws.Cells(i - 1, 6).value & "." & ws.Cells(i - 1, 7).value 'F열, G열
                nameStr = Replace(nameStr, " ", ".") ' 공백 제거
                nameStr = Replace(nameStr, "-", ".") ' 특수문자 대체
                
                ' 범위 설정
                Set rng = ws.Range("H" & startRow & ":H" & (i - 1)) ' H열
                
                ' 이름 정의
                ThisWorkbook.Names.Add Name:=nameStr, RefersTo:=rng
            End If
            
            ' 새 그룹 시작
            startRow = i
        End If
    Next i
    
    MsgBox "모든 범위에 이름이 자동으로 부여되었습니다!", vbInformation, "완료"
End Sub


'<기능>
'부여된 이름이 너무 많은 경우, 이름 상자(수식창 왼쪽) 및 이름관리자에서 너무 많은 내용이 표시되어 불편할 수 있는 문제가 있음.
'이름을 숨김처리하면 (1)이름상자에서 표시되지 않음 (2)이름관리자에서 표시되지 않음 (3)이름을 아는 경우에 한하여, 수식으로 불러오기 가능함.
'이름을 보임처리하면 (1)이름상자에서 표시됨 (2)이름관리자에서 표시되어 삭제할 수 있음 (3)수식으로 불러오기 가능함.

Sub ShowAllNames()
    Dim n As Name
    For Each n In ThisWorkbook.Names
        n.Visible = True
    Next n
    MsgBox "모든 숨겨진 이름이 다시 보이도록 설정되었습니다."
End Sub

Sub HideNames()
    Dim n As Name
    For Each n In ThisWorkbook.Names
        If InStr(1, n.Name, "List1") > 0 Or _
           InStr(1, n.Name, "List2_") > 0 Or _
           InStr(1, n.Name, "List3_") > 0 Or _
           InStr(1, n.Name, "Tbl_Source") > 0 Or _
           InStr(1, n.Name, "법원경매정보_") > 0 Or _
           InStr(1, n.Name, "실거래가_") > 0 Or _
           InStr(1, n.Name, "인포케어_") > 0 Or _
           InStr(1, n.Name, "KTest") > 0 Or _
           InStr(1, n.Name, "용도_") > 0 Then
            n.Visible = False 'True:보이기, False: 숨기기
        End If
    Next n
    MsgBox "이름이 숨김처리되었습니다. 숨겨진 이름은 파워쿼리에서 사용할 수 없습니다."
End Sub


