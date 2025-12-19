Attribute VB_Name = "Module_JM"
Sub SortAndMapCollateral()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim collateralIndex As Long
    Dim currentUsage As String
    Dim previousUsage As String
    
    ' Input 시트 설정
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' 데이터가 있는 마지막 행 찾기 (A열을 기준으로)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 용도(소분류) 컬럼을 기준으로 정렬 (예: U 컬럼이 '용도(소분류)')
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("U7:U" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A6:BV" & lastRow)  ' A6부터 BV까지의 범위 정렬
        .header = xlYes
        .Apply
    End With
    
    ' 담보_mapping 값 채우기 (AB 컬럼이 '담보_mapping' 컬럼)
    collateralIndex = 1
    previousUsage = ws.Range("U7").value  ' 첫 번째 '용도(소분류)' 값
    ws.Range("AB7").value = "담보-" & collateralIndex  ' 첫 번째 담보 값 설정

    For i = 8 To lastRow  ' 7행 이후로 데이터 작업
        currentUsage = ws.Range("U" & i).value
        
        If currentUsage = previousUsage Then
            ' '용도(소분류)' 값이 동일할 때
            ws.Range("AB" & i).value = "담보-" & collateralIndex
        Else
            ' '용도(소분류)' 값이 달라질 때
            collateralIndex = collateralIndex + 1
            ws.Range("AB" & i).value = "담보-" & collateralIndex
        End If
        
        previousUsage = currentUsage
    Next i
    
    ' 1. 커서를 AB5셀로 옮기고 화면 전환
    ws.Activate
    ws.Range("AB5").Select
    
    ' 2. AB6셀에 '담보_Mapping' 입력 및 서식 적용
    With ws.Range("AB6")
        .value = "담보_Mapping"
        .Font.Bold = True
        .Interior.Color = RGB(181, 230, 162)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 3. 열 너비를 14로 설정
    ws.columns("AB").ColumnWidth = 14
    
    ' 4. 메시지박스로 확인 요청
    MsgBox "담보가 자동으로 Mapping되었습니다." & vbCrLf & "Input' 시트의 '담보_Mapping' 열(AB열)에서" & vbCrLf & "물건의 담보 Mapping을 꼭 확인해주세요.", vbInformation + vbOKOnly, "확인"
    
End Sub

Sub LookupAndFillDataBySingleSheet_old(inputSheetName As String, lookupSheetName As String)
    Dim wsInput As Worksheet
    Dim wsLookup As Worksheet
    Dim lastRowInput As Long
    Dim lastRowLookup As Long
    Dim inputKey As String, lookupKey As String
    Dim i As Long, j As Long, lookupRow As Long
    Dim inputCol As Range
    Dim headerRow As Range
    Dim lookupCol As Range
    Dim matched As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets(inputSheetName)
    Set wsLookup = ThisWorkbook.Sheets(lookupSheetName)
    
    ' Input 시트와 Lookup 시트의 마지막 행 찾기
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "C").End(xlUp).row
    lastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, "A").End(xlUp).row
    
    ' Input 시트의 헤더 행 범위
    Set headerRow = wsInput.Rows(6)
    
    ' Input 시트의 모든 행에 대해 처리 (C열: 등기부등본고유번호)
    For i = 7 To lastRowInput
        inputKey = CleanString(wsInput.Cells(i, 3).value) ' C열에서 등기부등본고유번호 가져옴
        matched = False ' 매칭 여부 초기화
        
        ' Lookup 시트의 A열과 Input 시트의 C열 값 비교
        For lookupRow = 2 To lastRowLookup
            lookupKey = CleanString(wsLookup.Cells(lookupRow, 1).value) ' Lookup 시트의 A열 값
            
            ' 디버그: 비교하는 값들을 출력
            Debug.Print "Comparing: Input Key (" & inputKey & ") with Lookup Key (" & lookupKey & ")"
            
            If lookupKey = inputKey Then
                ' Key가 일치할 경우 헤더를 비교하며 값을 채움
                matched = True
                Debug.Print "Match found: " & inputKey & " at lookup row: " & lookupRow
                
                For j = 1 To wsLookup.Cells(1, wsLookup.columns.Count).End(xlToLeft).Column
                    Set lookupCol = wsLookup.Cells(1, j)
                    
                    If Not IsEmpty(lookupCol.value) Then
                        ' 헤더가 일치하는 경우 값을 채움 (헤더 비교를 CleanHeader로 정리)
                        Set inputCol = headerRow.Find(What:=CleanHeader(lookupCol.value), LookIn:=xlValues, LookAt:=xlWhole)
                        
                        If Not inputCol Is Nothing Then
                            wsInput.Cells(i, inputCol.Column).value = wsLookup.Cells(lookupRow, j).value
                            Debug.Print "Value copied: " & wsLookup.Cells(lookupRow, j).value & " to Input row: " & i & ", column: " & inputCol.Column
                        Else
                            Debug.Print "Matching header not found for: " & lookupCol.value
                        End If
                    End If
                Next j
                Exit For ' 매칭이 되면 다음 Input row로 넘어감
            End If
        Next lookupRow
        
        If Not matched Then
            ' 디버그: 매칭 실패 시 입력 키를 출력
            Debug.Print "Key not found: " & inputKey
        End If
    Next i

End Sub

Function CleanString(value As String) As String
    ' 모든 공백과 특수문자를 제거하고 문자열로 반환
    CleanString = Replace(Replace(Replace(Trim(CStr(value)), vbCrLf, ""), vbTab, ""), Chr(160), "")
End Function

Function CleanHeader(header As String) As String
    ' 모든 공백과 줄바꿈을 제거하고 헤더로 반환
    CleanHeader = Replace(Replace(Trim(CStr(header)), vbCrLf, ""), vbTab, "")
End Function

Sub LookupAndFillDataBySingleSheet(inputSheetName As String, lookupSheetName As String)
    Dim wsInput As Worksheet
    Dim wsLookup As Worksheet
    Dim lastRowInput As Long
    Dim lastRowLookup As Long
    Dim inputKey As String, lookupKey As String
    Dim i As Long, j As Long, lookupRow As Long
    Dim inputCol As Range
    Dim headerRow As Range
    Dim lookupCol As Range
    Dim matched As Boolean
    
    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets(inputSheetName)
    Set wsLookup = ThisWorkbook.Sheets(lookupSheetName)
    
    ' Input 시트와 Lookup 시트의 마지막 행 찾기
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "C").End(xlUp).row
    lastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, "A").End(xlUp).row
    
    ' Input 시트의 헤더 행 범위
    Set headerRow = wsInput.Rows(6)
    
    ' Input 시트의 모든 행에 대해 처리 (C열: 등기부등본고유번호)
    For i = 7 To lastRowInput
        inputKey = CleanString(wsInput.Cells(i, 3).value) ' C열에서 등기부등본고유번호 가져옴
        matched = False ' 매칭 여부 초기화
        
        ' Lookup 시트의 A열과 Input 시트의 C열 값 비교
        For lookupRow = 2 To lastRowLookup
            lookupKey = CleanString(wsLookup.Cells(lookupRow, 1).value) ' Lookup 시트의 A열 값
            
            ' 디버그: 비교하는 값들을 출력
            Debug.Print "Comparing: Input Key (" & inputKey & ") with Lookup Key (" & lookupKey & ")"
            
            If lookupKey = inputKey Then
                ' Key가 일치할 경우 헤더를 비교하며 값을 채움
                matched = True
                Debug.Print "Match found: " & inputKey & " at lookup row: " & lookupRow
                
                For j = 1 To wsLookup.Cells(1, wsLookup.columns.Count).End(xlToLeft).Column
                    Set lookupCol = wsLookup.Cells(1, j)
                    
                    If Not IsEmpty(lookupCol.value) Then
                        ' 헤더가 일치하는 경우 값을 채움 (헤더 비교를 CleanHeader로 정리)
                        Set inputCol = headerRow.Find(What:=CleanHeader(lookupCol.value), LookIn:=xlValues, LookAt:=xlWhole)
                        
                        If Not inputCol Is Nothing Then
                            wsInput.Cells(i, inputCol.Column).value = wsLookup.Cells(lookupRow, j).value
                            Debug.Print "Value copied: " & wsLookup.Cells(lookupRow, j).value & " to Input row: " & i & ", column: " & inputCol.Column
                        Else
                            Debug.Print "Matching header not found for: " & lookupCol.value
                        End If
                    End If
                Next j
                Exit For ' 매칭이 되면 다음 Input row로 넘어감
            End If
        Next lookupRow
        
        If Not matched Then
            ' 디버그: 매칭 실패 시 입력 키를 출력
            Debug.Print "Key not found: " & inputKey
        End If
    Next i
    
    ' 추가된 부분: 담보_mapping 열을 유지하면서 토지_공시지가_2024 기준으로 내림차순 정렬
    Call SortByCollateralAndLandPrice(inputSheetName)
End Sub

Sub SortByCollateralAndLandPrice(inputSheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' 시트 설정
    Set ws = ThisWorkbook.Sheets(inputSheetName)
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 담보_mapping 및 토지_공시지가_2024 기준 정렬
    With ws.Sort
        .SortFields.Clear
        ' 담보_mapping 열(AB 열)을 기준으로 먼저 오름차순 정렬
        .SortFields.Add key:=ws.Range("AB7:AB" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' 토지_공시지가_2024 열(AN 열)을 기준으로 내림차순 정렬
        .SortFields.Add key:=ws.Range("AN7:AN" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("A6:BV" & lastRow)  ' 정렬할 범위 설정
        .header = xlYes  ' 헤더 포함
        .Apply
    End With
End Sub
Sub 담보물정보table_생성()
    Dim wsInput As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowInput As Long
    Dim uniqueCollaterals As Collection
    Dim rng As Range
    Dim cell As Range
    Dim collateral As Variant
    Dim jibun As Variant
    Dim firstRow As Long
    Dim currentRow As Long
    Dim tbl As ListObject
    Dim copyRange As Range
    Dim newTable As ListObject
    Dim firstUsageValue As String
    Dim buildingSum As Double
    Dim buildingStructure As String
    Dim landAreaSum As Double
    Dim filterRange As Range
    Dim visibleCells As Range
    Dim buildingValue As Double
    Dim uniqueBuildingStructures As Collection
    Dim uniqueLandUsages As Collection
    Dim uniqueBuildingReceptionDates As Collection
    Dim structValue As Variant
    Dim landUsageValue As Variant
    Dim receptionDateValue As Variant
    Dim buildingReceptionDate As String
    Dim usageType As String
    Dim areaValue As Variant
    Dim jibunValue As String
    Dim landPrice2024 As Variant
    Dim landPrice2023 As Variant
    Dim currentTableRow As Long
    Dim jibunRowDict As Object
    Dim foundRow As Long
    Dim rowCell As Range
    Dim landPriceFormula As String
    Dim UsageTypeCol As Long
    Dim DaejiAreaCol As Long
    Dim landAreaCol As Long
    Dim LandUsageCol As Long
    Dim BuildingAreaCol As Long
    Dim BuildingStructCol As Long
    Dim BuildingReceptionDateCol As Long
    Dim JibunCol As Long
    Dim LandPrice2024Col As Long
    Dim LandPrice2023Col As Long
    Dim prevAutoFillFormulasInLists As Boolean
    Dim tblRange As Range
    Dim tableStartColumn As Long

    ' 시트 설정
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsTemplate = ThisWorkbook.Sheets("Tpl_담보물정보")
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsInput) ' 새로운 시트에 복사
    wsOutput.Name = "담보물정보"

    ' Input 시트의 마지막 행 찾기 (A열 기준)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).row

    ' '담보_mapping' 열에서 고유한 담보 값 추출
    Set uniqueCollaterals = New Collection
    On Error Resume Next
    For Each cell In wsInput.Range("AB7:AB" & lastRowInput)
        collateral = cell.value
        If collateral <> "" Then
            uniqueCollaterals.Add collateral, CStr(collateral) ' 고유한 값만 추가
        End If
    Next cell
    On Error GoTo 0

    ' 'Tpl_담보물정보' 시트의 '담보물정보' 표 복사
    Set tbl = wsTemplate.ListObjects("담보물정보")
    Set copyRange = tbl.Range

    ' 필요한 컬럼 인덱스 계산 (고정된 값)
    UsageTypeCol = 4        ' D열
    DaejiAreaCol = 35       ' AI열
    landAreaCol = 29        ' AC열
    LandUsageCol = 38       ' AL열
    BuildingAreaCol = 32    ' AF열
    BuildingStructCol = 31  ' AE열
    BuildingReceptionDateCol = 36 ' AJ열
    JibunCol = 37           ' AK열
    LandPrice2024Col = 40   ' AN열
    LandPrice2023Col = 41   ' AO열

    ' 새로운 시트에 담보별로 '표' 복사 및 값 채우기
    currentRow = 1
    For Each collateral In uniqueCollaterals
        ' '담보_mapping' 값에 해당하는 행 필터링 (Input 시트)
        wsInput.Range("A6:BV" & lastRowInput).AutoFilter Field:=28, Criteria1:=collateral ' AB열(28번째 컬럼)에 해당하는 담보_mapping 값 기준 필터링

        ' 필터링된 데이터 중 첫 번째 데이터 행 추출
        On Error Resume Next
        Set visibleCells = wsInput.Range("A7:BV" & lastRowInput).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not visibleCells Is Nothing Then
            ' 'Tpl_담보물정보' 시트의 표 복사 후 새로운 시트에 붙여넣기
            copyRange.Copy
            wsOutput.Cells(currentRow, 1).PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False

            ' 새로운 표에 값 채우기
            Set newTable = wsOutput.ListObjects(wsOutput.ListObjects.Count)
            firstRow = newTable.HeaderRowRange.row + 1
            tableStartColumn = newTable.HeaderRowRange.Column ' 테이블의 시작 열 번호

            ' '토지_용도지역명_1' 값 고유하게 수집 (AL열: 38번째 열)
            Set uniqueLandUsages = New Collection
            On Error Resume Next
            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    landUsageValue = wsInput.Cells(rowCell.row, LandUsageCol).value
                    If landUsageValue <> "" Then
                        uniqueLandUsages.Add landUsageValue, CStr(landUsageValue) ' 고유한 값만 추가
                    End If
                End If
            Next rowCell
            On Error GoTo 0

            ' 고유한 '토지_용도지역명_1' 값을 " / "로 연결
            firstUsageValue = ""
            For Each landUsageValue In uniqueLandUsages
                If firstUsageValue = "" Then
                    firstUsageValue = landUsageValue
                Else
                    firstUsageValue = firstUsageValue & " / " & landUsageValue
                End If
            Next landUsageValue

            ' '건물내역' 값 합산 (AF열: 32번째 열)
            buildingSum = 0 ' 초기화
            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    buildingValue = wsInput.Cells(rowCell.row, BuildingAreaCol).value
                    If IsNumeric(buildingValue) Then
                        buildingSum = buildingSum + buildingValue
                    End If
                End If
            Next rowCell

            ' '건물구조' 값 고유하게 수집 (AE열: 31번째 열)
            Set uniqueBuildingStructures = New Collection
            On Error Resume Next
            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    structValue = wsInput.Cells(rowCell.row, BuildingStructCol).value
                    If structValue <> "" Then
                        uniqueBuildingStructures.Add structValue, CStr(structValue) ' 고유한 값만 추가
                    End If
                End If
            Next rowCell
            On Error GoTo 0

            ' 고유한 '건물구조' 값을 " / "로 연결
            buildingStructure = ""
            For Each structValue In uniqueBuildingStructures
                If buildingStructure = "" Then
                    buildingStructure = structValue
                Else
                    buildingStructure = buildingStructure & " / " & structValue
                End If
            Next structValue

            ' '건물_접수일' 값 고유하게 수집 (AJ열: 36번째 열)
            Set uniqueBuildingReceptionDates = New Collection
            On Error Resume Next
            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    receptionDateValue = wsInput.Cells(rowCell.row, BuildingReceptionDateCol).value
                    If receptionDateValue <> "" Then
                        uniqueBuildingReceptionDates.Add receptionDateValue, CStr(receptionDateValue) ' 고유한 값만 추가
                    End If
                End If
            Next rowCell
            On Error GoTo 0

            ' 고유한 '건물_접수일' 값을 " / "로 연결
            buildingReceptionDate = ""
            If uniqueBuildingReceptionDates.Count > 0 Then
                For Each receptionDateValue In uniqueBuildingReceptionDates
                    If buildingReceptionDate = "" Then
                        buildingReceptionDate = receptionDateValue
                    Else
                        buildingReceptionDate = buildingReceptionDate & " / " & receptionDateValue
                    End If
                Next receptionDateValue

                ' 표의 첫 번째 행의 '8' 열에 값 채우기
                wsOutput.Cells(firstRow, newTable.ListColumns("8").Range.Column).value = buildingReceptionDate
            End If

            ' '토지_면적' 또는 '대지권_대상_면적' 값 합산
            landAreaSum = 0 ' 초기화
            ' 수식 생성 초기화
            landPriceFormula = "="

            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    usageType = wsInput.Cells(rowCell.row, UsageTypeCol).value

                    ' '토지_공시지가_2024' 값 가져오기
                    landPrice2024 = wsInput.Cells(rowCell.row, LandPrice2024Col).value
                    If Not IsNumeric(landPrice2024) Or IsEmpty(landPrice2024) Then
                        landPrice2024 = 0
                    End If

                    ' 면적 값 가져오기
                    If usageType = "집합건물" Then
                        areaValue = wsInput.Cells(rowCell.row, DaejiAreaCol).value
                    Else
                        areaValue = wsInput.Cells(rowCell.row, landAreaCol).value
                    End If
                    If Not IsNumeric(areaValue) Or IsEmpty(areaValue) Then
                        areaValue = 0
                    End If

                    ' 면적 합산
                    landAreaSum = landAreaSum + areaValue

                    ' 수식 생성
                    If landPriceFormula = "=" Then
                        landPriceFormula = landPriceFormula & areaValue & "*" & landPrice2024
                    Else
                        landPriceFormula = landPriceFormula & "+" & areaValue & "*" & landPrice2024
                    End If
                End If
            Next rowCell

            ' 표의 첫 번째 행에 값 채우기
            wsOutput.Cells(firstRow, newTable.ListColumns("1").Index).value = firstUsageValue
            wsOutput.Cells(firstRow + 1, newTable.ListColumns("1").Index).value = landAreaSum
            wsOutput.Cells(firstRow, newTable.ListColumns("5").Index).value = buildingSum
            wsOutput.Cells(firstRow + 2, newTable.ListColumns("5").Index).value = buildingStructure

            ' 자동 채우기 기능 비활성화
            prevAutoFillFormulasInLists = Application.AutoCorrect.AutoFillFormulasInLists
            Application.AutoCorrect.AutoFillFormulasInLists = False

            ' '표'의 '4' 열의 3번째 행에 수식 입력 (반복문 밖에서 수행)
            wsOutput.Cells(firstRow + 2, newTable.ListColumns("4").Index).Formula = landPriceFormula

            ' 자동 채우기 기능 원래대로 복원
            Application.AutoCorrect.AutoFillFormulasInLists = prevAutoFillFormulasInLists

            ' 새로운 행에 값을 추가해야 함
            currentTableRow = firstRow + 3 ' 기본 행 위치

            ' '지번' 값 고유하게 수집 및 행 번호 저장 (AK열: 37번째 열)
            Set jibunRowDict = CreateObject("Scripting.Dictionary")
            On Error Resume Next
            For Each rowCell In visibleCells.columns(1).Cells
                If rowCell.row > 6 Then
                    jibunValue = wsInput.Cells(rowCell.row, JibunCol).value
                    If jibunValue <> "" Then
                        If Not jibunRowDict.Exists(jibunValue) Then
                            jibunRowDict.Add jibunValue, rowCell.row ' wsInput에서의 행 번호 저장
                        End If
                    End If
                End If
            Next rowCell
            On Error GoTo 0

            ' 추가로 필요한 행 수 계산
            Dim initialJibunRows As Integer
            initialJibunRows = 2 ' 초기 테이블에 '지번' 데이터를 위한 행 수

            Dim additionalRowsNeeded As Integer
            additionalRowsNeeded = (jibunRowDict.Count * 2) - initialJibunRows

            If additionalRowsNeeded > 0 Then
                ' 테이블에 필요한 행 추가
                Dim insertPosition As Long
                insertPosition = currentTableRow - newTable.HeaderRowRange.row + 1 ' 테이블 내 위치 계산

                For i = 1 To additionalRowsNeeded
                    newTable.ListRows.Add Position:=insertPosition
                Next i
            End If

            ' '지번' 값마다 데이터 입력
            For Each jibun In jibunRowDict.Keys
                ' 행에 '지번' 값 채우기
                wsOutput.Cells(currentTableRow, newTable.ListColumns("1").Index).value = jibun
                wsOutput.Cells(currentTableRow + 1, newTable.ListColumns("1").Index).value = jibun

                ' 딕셔너리에서 행 번호 가져오기
                foundRow = jibunRowDict(jibun)

                ' '토지_공시지가_2024', '토지_공시지가_2023' 값 가져오기
                landPrice2024 = wsInput.Cells(foundRow, LandPrice2024Col).value
                landPrice2023 = wsInput.Cells(foundRow, LandPrice2023Col).value

                If Not IsNumeric(landPrice2024) Or IsEmpty(landPrice2024) Then
                    landPrice2024 = 0
                End If
                If Not IsNumeric(landPrice2023) Or IsEmpty(landPrice2023) Then
                    landPrice2023 = 0
                End If

                wsOutput.Cells(currentTableRow, newTable.ListColumns("2").Index).value = landPrice2024
                wsOutput.Cells(currentTableRow + 1, newTable.ListColumns("2").Index).value = landPrice2023

                ' '(1) 토지' 컬럼에 날짜 값 채우기
                wsOutput.Cells(currentTableRow, newTable.ListColumns("(1) 토지").Index).value = "2024-01-01"
                wsOutput.Cells(currentTableRow + 1, newTable.ListColumns("(1) 토지").Index).value = "2023-01-01"

                ' 다음 '지번'을 위한 행 이동
                currentTableRow = currentTableRow + 2
            Next jibun

            ' 현재 표 범위 저장
            Set tblRange = newTable.Range

            ' 해당 표의 '(1) 토지' 열을 내림차순으로 정렬
            Call SortLandColumn(tblRange)

            ' 다음 표로 이동
            currentRow = tblRange.Rows.Count + currentRow + 2

        End If

        ' 필터 해제
        wsInput.AutoFilterMode = False

        ' 다음 표로 이동
        currentRow = currentTableRow + 2
    Next collateral
End Sub

Sub SortLandColumn(tblRange As Range)
    ' 표 내에서 '(1) 토지' 열 (첫 번째 열) 기준으로 내림차순 정렬
    With tblRange.Worksheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=tblRange.columns(1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange tblRange
        .header = xlYes
        .Apply
    End With
End Sub
Sub 실거래사례_국토(sheetName As String)
    Dim wsInput As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowInput As Long
    Dim tbl As ListObject
    Dim copyRange As Range
    Dim newTable As ListObject
    Dim visibleCells As Range
    Dim dealDateCol As Long, addrCol As Long, excluUseArCol As Long, floorCol As Long, dealAmountCol As Long
    Dim reportCol As Range
    Dim additionalRowsNeeded As Long
    Dim newRow As ListRow
    Dim filteredRowCount As Long
    Dim i As Long

    ' 입력받은 시트 설정
    Set wsInput = ThisWorkbook.Sheets(sheetName)
    ' 사전 정의된 테이블이 있는 템플릿 시트 설정
    Set wsTemplate = ThisWorkbook.Sheets("Tpl_실거래사례_국토")

    ' 새 시트 생성
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOutput.Name = "실거래사례_국토"

    ' Input 시트의 마지막 행 찾기 (A열 기준)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).row

    ' "보고서포함_여부" 열 찾기 (1번째 행에서)
    Set reportCol = wsInput.Rows(1).Find("보고서포함_여부")
    
    If reportCol Is Nothing Then
        MsgBox "'보고서포함_여부' 열을 찾을 수 없습니다."
        Exit Sub
    End If

    ' "보고서포함_여부" 열에서 'V' 값 필터링
    wsInput.Range("A1:Z" & lastRowInput).AutoFilter Field:=reportCol.Column, Criteria1:="V"

    ' 필터링된 행을 가져오기
    On Error Resume Next
    Set visibleCells = wsInput.Range("A2:Z" & lastRowInput).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' 필터링된 데이터가 없으면 종료
    If visibleCells Is Nothing Then
        MsgBox "필터링된 데이터가 없습니다."
        Exit Sub
    End If

    ' 필터링된 row의 수를 계산
    filteredRowCount = visibleCells.Rows.Count

    ' 'Tpl_실거래사례' 시트의 '실거래사례' 테이블 복사
    Set tbl = wsTemplate.ListObjects("실거래사례_국토")
    Set copyRange = tbl.Range
    copyRange.Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    Set newTable = wsOutput.ListObjects(1)

    ' 추가해야 할 row 계산 (필터된 데이터가 2줄 이상일 때만 추가)
    If filteredRowCount > 2 Then
        additionalRowsNeeded = filteredRowCount - 2
        ' 필요한 만큼 행을 추가
        For i = 1 To additionalRowsNeeded
            newTable.ListRows.Add
        Next i
    End If

    ' 필요한 열 인덱스 찾기 (1번째 행에서)
    dealDateCol = wsInput.Rows(1).Find("dealDate").Column
    addrCol = wsInput.Rows(1).Find("대상주소").Column
    excluUseArCol = wsInput.Rows(1).Find("excluUseAr").Column
    floorCol = wsInput.Rows(1).Find("floor").Column
    dealAmountCol = wsInput.Rows(1).Find("dealAmount").Column

    ' 필터링된 데이터 입력
    i = 1
    For Each row In visibleCells.Rows
        If row.row > 1 Then ' 데이터 행만 처리
            ' 먼저 사전 정의된 2개의 행을 채운다
            If i <= 2 Then
                With newTable.ListRows(i)
                    .Range(1, newTable.ListColumns("순번").Index).value = i
                    .Range(1, newTable.ListColumns("거래일자").Index).value = wsInput.Cells(row.row, dealDateCol).value
                    .Range(1, newTable.ListColumns("주소").Index).value = wsInput.Cells(row.row, addrCol).value
                    .Range(1, newTable.ListColumns("전용면적(㎡)").Index).value = wsInput.Cells(row.row, excluUseArCol).value
                    .Range(1, newTable.ListColumns("층수").Index).value = wsInput.Cells(row.row, floorCol).value
                    .Range(1, newTable.ListColumns("거래금액").Index).value = wsInput.Cells(row.row, dealAmountCol).value * 10000 ' 10,000 곱하기
                End With
            Else
                ' 이후 필요한 행을 추가하고 채운다
                Set newRow = newTable.ListRows.Add
                newRow.Range(1, newTable.ListColumns("순번").Index).value = i
                newRow.Range(1, newTable.ListColumns("거래일자").Index).value = wsInput.Cells(row.row, dealDateCol).value
                newRow.Range(1, newTable.ListColumns("주소").Index).value = wsInput.Cells(row.row, addrCol).value
                newRow.Range(1, newTable.ListColumns("전용면적(㎡)").Index).value = wsInput.Cells(row.row, excluUseArCol).value
                newRow.Range(1, newTable.ListColumns("층수").Index).value = wsInput.Cells(row.row, floorCol).value
                newRow.Range(1, newTable.ListColumns("거래금액").Index).value = wsInput.Cells(row.row, dealAmountCol).value * 10000 ' 10,000 곱하기
            End If

            i = i + 1
        End If
    Next row

    ' 필터 해제
    wsInput.AutoFilterMode = False
End Sub
Sub 실거래사례_밸류맵(sheetName As String)
    Dim wsInput As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowInput As Long
    Dim tbl As ListObject
    Dim copyRange As Range
    Dim newTable As ListObject
    Dim visibleCells As Range
    Dim dealDateCol As Long, addrCol As Long, excluUseArCol As Long, landAreaCol As Long, buildingTypeCol As Long, dealAmountCol As Long, floorCol As Long
    Dim reportCol As Range
    Dim additionalRowsNeeded As Long
    Dim newRow As ListRow
    Dim filteredRowCount As Long
    Dim i As Long
    Dim dealYearMonth As String

    ' 입력받은 시트 설정
    Set wsInput = ThisWorkbook.Sheets(sheetName)
    ' 사전 정의된 테이블이 있는 템플릿 시트 설정
    Set wsTemplate = ThisWorkbook.Sheets("Tpl_실거래사례_밸류맵")

    ' 새 시트 생성
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOutput.Name = "실거래사례_밸류맵"

    ' Input 시트의 마지막 행 찾기 (A열 기준)
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).row

    ' "보고서포함_여부" 열 찾기 (1번째 행에서)
    Set reportCol = wsInput.Rows(1).Find("보고서포함_여부")
    
    If reportCol Is Nothing Then
        MsgBox "'보고서포함_여부' 열을 찾을 수 없습니다."
        Exit Sub
    End If

    ' "보고서포함_여부" 열에서 'V' 값 필터링
    wsInput.Range("A1:Z" & lastRowInput).AutoFilter Field:=reportCol.Column, Criteria1:="V"

    ' 필터링된 행을 가져오기
    On Error Resume Next
    Set visibleCells = wsInput.Range("A2:Z" & lastRowInput).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' 필터링된 데이터가 없으면 종료
    If visibleCells Is Nothing Then
        MsgBox "필터링된 데이터가 없습니다."
        Exit Sub
    End If

    ' 필터링된 row의 수를 계산
    filteredRowCount = visibleCells.Rows.Count

    ' 'Tpl_실거래사례_밸류맵' 시트의 '실거래사례_밸류맵' 테이블 복사
    Set tbl = wsTemplate.ListObjects("실거래사례_밸류맵")
    Set copyRange = tbl.Range
    copyRange.Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    Set newTable = wsOutput.ListObjects(1)

    ' 추가해야 할 row 계산 (필터된 데이터가 2줄 이상일 때만 추가)
    If filteredRowCount > 2 Then
        additionalRowsNeeded = filteredRowCount - 2
        ' 필요한 만큼 행을 추가
        For i = 1 To additionalRowsNeeded
            newTable.ListRows.Add
        Next i
    End If

    ' 필요한 열 인덱스 찾기 (1번째 행에서)
    dealDateCol = wsInput.Rows(1).Find("거래년월").Column
    addrCol = wsInput.Rows(1).Find("주소").Column
    excluUseArCol = wsInput.Rows(1).Find("연(전유)면적").Column
    landAreaCol = wsInput.Rows(1).Find("토지면적").Column
    buildingTypeCol = wsInput.Rows(1).Find("지목").Column
    dealAmountCol = wsInput.Rows(1).Find("거래금액").Column
    floorCol = wsInput.Rows(1).Find("층(상가)").Column ' 층(상가) 열 인덱스 추가

    ' 필터링된 데이터 입력
    i = 1
    For Each row In visibleCells.Rows
        If row.row > 1 Then ' 데이터 행만 처리
            ' 먼저 사전 정의된 2개의 행을 채운다
            If i <= 2 Then
                With newTable.ListRows(i)
                    ' 순번 입력
                    .Range(1, newTable.ListColumns("순번").Index).value = i
                    ' 거래년월에서 '.'을 '-'로 변환
                    dealYearMonth = Replace(wsInput.Cells(row.row, dealDateCol).value, ".", "-")
                    .Range(1, newTable.ListColumns("거래일자").Index).value = dealYearMonth
                    .Range(1, newTable.ListColumns("주소").Index).value = wsInput.Cells(row.row, addrCol).value
                    .Range(1, newTable.ListColumns("건물연면적(㎡)").Index).value = wsInput.Cells(row.row, excluUseArCol).value
                    .Range(1, newTable.ListColumns("대지면적(㎡)").Index).value = wsInput.Cells(row.row, landAreaCol).value
                    .Range(1, newTable.ListColumns("지목").Index).value = wsInput.Cells(row.row, buildingTypeCol).value
                    .Range(1, newTable.ListColumns("거래금액").Index).value = wsInput.Cells(row.row, dealAmountCol).value
                    .Range(1, newTable.ListColumns("층수").Index).value = wsInput.Cells(row.row, floorCol).value ' 층수 매핑
                End With
            Else
                ' 이후 필요한 행을 추가하고 채운다
                Set newRow = newTable.ListRows.Add
                ' 순번 입력
                newRow.Range(1, newTable.ListColumns("순번").Index).value = i
                dealYearMonth = Replace(wsInput.Cells(row.row, dealDateCol).value, ".", "-")
                newRow.Range(1, newTable.ListColumns("거래일자").Index).value = dealYearMonth
                newRow.Range(1, newTable.ListColumns("주소").Index).value = wsInput.Cells(row.row, addrCol).value
                newRow.Range(1, newTable.ListColumns("건물연면적(㎡)").Index).value = wsInput.Cells(row.row, excluUseArCol).value
                newRow.Range(1, newTable.ListColumns("대지면적(㎡)").Index).value = wsInput.Cells(row.row, landAreaCol).value
                newRow.Range(1, newTable.ListColumns("지목").Index).value = wsInput.Cells(row.row, buildingTypeCol).value
                newRow.Range(1, newTable.ListColumns("거래금액").Index).value = wsInput.Cells(row.row, dealAmountCol).value
                newRow.Range(1, newTable.ListColumns("층수").Index).value = wsInput.Cells(row.row, floorCol).value ' 층수 매핑
            End If

            i = i + 1
        End If
    Next row

    ' 필터 해제
    wsInput.AutoFilterMode = False
End Sub

'여러 테이블이 존재하는 여러 시트에 대해 모두 범위로 반환하고 병합하는 함수
Sub 담보물정보_실거래사례_병합(sheetNames As Variant)
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowOutput As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim currentSheetName As Variant
    Dim currentRange As Range
    Dim pasteRow As Long
    Dim lastRow As Long
    Dim newSheetName As String
    Dim tableRange As Range
    Dim cell As Range
    
    ' 새로운 시트 생성
    newSheetName = "담보물정보_사례"
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(newSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = newSheetName
    Else
        wsOutput.Cells.Clear ' 기존 시트가 있으면 내용을 초기화
    End If
    On Error GoTo 0
    
    ' 데이터를 붙여넣을 첫 번째 행 설정
    pasteRow = 1

    ' 입력받은 시트들에 대해 반복
    For Each currentSheetName In sheetNames
        Set wsInput = ThisWorkbook.Sheets(currentSheetName)
        
        ' 해당 시트의 모든 테이블을 반복하여 범위로 변환 후 붙여넣기
        For Each tbl In wsInput.ListObjects
            ' 테이블을 범위로 변환
            Set tableRange = tbl.Range
            tbl.Unlist ' 테이블 해제
            
            ' 테이블 범위를 새로운 시트에 수식 포함하여 붙여넣기
            tableRange.Copy
            wsOutput.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteAll ' 수식 포함하여 붙여넣기
            Application.CutCopyMode = False
            
            ' 붙여넣은 후 2줄 간격으로 다음 범위 붙여넣기
            pasteRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row + 3
        Next tbl
    Next currentSheetName
    
    ' "텍스트 형식으로 저장된 숫자" 경고 없애기
    On Error Resume Next
    For Each cell In wsOutput.UsedRange
        If IsNumeric(cell.value) And cell.NumberFormat = "@" Then
            cell.NumberFormat = "0" ' 숫자 형식으로 변경
        End If
    Next cell

    ' 경고 제거 실행
    For Each cell In wsOutput.UsedRange
        If cell.Errors(xlNumberAsText).value = True Then
            cell.Errors(xlNumberAsText).Ignore = True ' 텍스트 형식 경고 무시
        End If
    Next cell
    On Error GoTo 0
End Sub
Sub 담보물정보_범위변환(sheetNames As Variant)
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowOutput As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim currentSheetName As Variant
    Dim currentRange As Range
    Dim pasteRow As Long
    Dim lastRow As Long
    Dim newSheetName As String
    Dim tableRange As Range
    Dim cell As Range
    
    ' 새로운 시트 생성
    newSheetName = "담보물정보_2"
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(newSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = newSheetName
    Else
        wsOutput.Cells.Clear ' 기존 시트가 있으면 내용을 초기화
    End If
    On Error GoTo 0
    
    ' 데이터를 붙여넣을 첫 번째 행 설정
    pasteRow = 1

    ' 입력받은 시트들에 대해 반복
    For Each currentSheetName In sheetNames
        Set wsInput = ThisWorkbook.Sheets(currentSheetName)
        
        ' 해당 시트의 모든 테이블을 반복하여 범위로 변환 후 붙여넣기
        For Each tbl In wsInput.ListObjects
            ' 테이블을 범위로 변환
            Set tableRange = tbl.Range
            tbl.Unlist ' 테이블 해제
            
            ' 테이블 범위를 새로운 시트에 수식 포함하여 붙여넣기
            tableRange.Copy
            wsOutput.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteAll ' 수식 포함하여 붙여넣기
            Application.CutCopyMode = False
            
            ' 붙여넣은 후 2줄 간격으로 다음 범위 붙여넣기
            pasteRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row + 3
        Next tbl
    Next currentSheetName
    
    ' "텍스트 형식으로 저장된 숫자" 경고 없애기
    On Error Resume Next
    For Each cell In wsOutput.UsedRange
        If IsNumeric(cell.value) And cell.NumberFormat = "@" Then
            cell.NumberFormat = "0" ' 숫자 형식으로 변경
        End If
    Next cell

    ' 경고 제거 실행
    For Each cell In wsOutput.UsedRange
        If cell.Errors(xlNumberAsText).value = True Then
            cell.Errors(xlNumberAsText).Ignore = True ' 텍스트 형식 경고 무시
        End If
    Next cell
    On Error GoTo 0
    
    ' 기존 table이 존재하던 담보물정보 sheet 삭제
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("담보물정보").Delete
    Application.DisplayAlerts = True
    
    ' "담보물정보_2" 시트명을 "담보물정보"로 변경
    ThisWorkbook.Sheets("담보물정보_2").Name = "2-1"
End Sub
Sub 실거래사례_국토_범위변환(sheetNames As Variant)
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowOutput As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim currentSheetName As Variant
    Dim currentRange As Range
    Dim pasteRow As Long
    Dim lastRow As Long
    Dim newSheetName As String
    Dim tableRange As Range
    Dim cell As Range
    
    ' 새로운 시트 생성
    newSheetName = "실거래사례_국토_2"
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(newSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = newSheetName
    Else
        wsOutput.Cells.Clear ' 기존 시트가 있으면 내용을 초기화
    End If
    On Error GoTo 0
    
    ' 데이터를 붙여넣을 첫 번째 행 설정
    pasteRow = 1

    ' 입력받은 시트들에 대해 반복
    For Each currentSheetName In sheetNames
        Set wsInput = ThisWorkbook.Sheets(currentSheetName)
        
        ' 해당 시트의 모든 테이블을 반복하여 범위로 변환 후 붙여넣기
        For Each tbl In wsInput.ListObjects
            ' 테이블을 범위로 변환
            Set tableRange = tbl.Range
            tbl.Unlist ' 테이블 해제
            
            ' 테이블 범위를 새로운 시트에 수식 포함하여 붙여넣기
            tableRange.Copy
            wsOutput.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteAll ' 수식 포함하여 붙여넣기
            Application.CutCopyMode = False
            
            ' 붙여넣은 후 2줄 간격으로 다음 범위 붙여넣기
            pasteRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row + 3
        Next tbl
    Next currentSheetName
    
    ' "텍스트 형식으로 저장된 숫자" 경고 없애기
    On Error Resume Next
    For Each cell In wsOutput.UsedRange
        If IsNumeric(cell.value) And cell.NumberFormat = "@" Then
            cell.NumberFormat = "0" ' 숫자 형식으로 변경
        End If
    Next cell

    ' 경고 제거 실행
    For Each cell In wsOutput.UsedRange
        If cell.Errors(xlNumberAsText).value = True Then
            cell.Errors(xlNumberAsText).Ignore = True ' 텍스트 형식 경고 무시
        End If
    Next cell
    On Error GoTo 0
    
    ' 기존 table이 존재하던 실거래사례_국토 sheet 삭제
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("실거래사례_국토").Delete
    Application.DisplayAlerts = True
    
    ' "실거래사례_국토_2" 시트명을 "실거래사례_국토"로 변경
    ThisWorkbook.Sheets("실거래사례_국토_2").Name = "6-1"
End Sub
Sub 실거래사례_밸류맵_범위변환(sheetNames As Variant)
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowOutput As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim currentSheetName As Variant
    Dim currentRange As Range
    Dim pasteRow As Long
    Dim lastRow As Long
    Dim newSheetName As String
    Dim tableRange As Range
    Dim cell As Range
    
    ' 새로운 시트 생성
    newSheetName = "실거래사례_밸류맵_2"
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets(newSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = newSheetName
    Else
        wsOutput.Cells.Clear ' 기존 시트가 있으면 내용을 초기화
    End If
    On Error GoTo 0
    
    ' 데이터를 붙여넣을 첫 번째 행 설정
    pasteRow = 1

    ' 입력받은 시트들에 대해 반복
    For Each currentSheetName In sheetNames
        Set wsInput = ThisWorkbook.Sheets(currentSheetName)
        
        ' 해당 시트의 모든 테이블을 반복하여 범위로 변환 후 붙여넣기
        For Each tbl In wsInput.ListObjects
            ' 테이블을 범위로 변환
            Set tableRange = tbl.Range
            tbl.Unlist ' 테이블 해제
            
            ' 테이블 범위를 새로운 시트에 수식 포함하여 붙여넣기
            tableRange.Copy
            wsOutput.Cells(pasteRow, 1).PasteSpecial Paste:=xlPasteAll ' 수식 포함하여 붙여넣기
            Application.CutCopyMode = False
            
            ' 붙여넣은 후 2줄 간격으로 다음 범위 붙여넣기
            pasteRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row + 3
        Next tbl
    Next currentSheetName
    
    ' "텍스트 형식으로 저장된 숫자" 경고 없애기
    On Error Resume Next
    For Each cell In wsOutput.UsedRange
        If IsNumeric(cell.value) And cell.NumberFormat = "@" Then
            cell.NumberFormat = "0" ' 숫자 형식으로 변경
        End If
    Next cell

    ' 경고 제거 실행
    For Each cell In wsOutput.UsedRange
        If cell.Errors(xlNumberAsText).value = True Then
            cell.Errors(xlNumberAsText).Ignore = True ' 텍스트 형식 경고 무시
        End If
    Next cell
    On Error GoTo 0
    
    ' 기존 table이 존재하던 실거래사례_밸류맵 sheet 삭제
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("실거래사례_밸류맵").Delete
    Application.DisplayAlerts = True
    
    ' "실거래사례_국토_2" 시트명을 "실거래사례_밸류맵"로 변경
    ThisWorkbook.Sheets("실거래사례_밸류맵_2").Name = "6-2"
End Sub
Sub 서식동기화(sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim formatRange As Range

    ' 지정된 시트 참조
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 데이터가 있는 마지막 행 계산 (A열 기준)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' 데이터가 있는 마지막 열 계산 (2행 기준)
    lastCol = ws.Cells(2, ws.columns.Count).End(xlToLeft).Column

    ' 2번 행(A2부터 마지막 열까지)의 서식 복사
    Set formatRange = ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
    formatRange.Copy

    ' 복사된 서식을 마지막 행까지 붙여넣기
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' 테이블 범위 외곽 테두리 적용
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Sub 담보물정보_생성() ' 버튼명 : 담보 Mapping
    ' 담보-x 생성
    Call SortAndMapCollateral
    ' 담보-x 생성 후 특정 담보물 수동 mapping 필요
End Sub
Sub 담보물정보_table_생성() ' 버튼명 : 담보물정보
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("2-1")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call 담보물정보_headersetting
    Call LookupAndFillDataBySingleSheet("Input", "Output_등본조회")
    Call LookupAndFillDataBySingleSheet("Input", "Output_공시지가")
    Call ApplyAllBorders_FromA6("Input")
    Call 담보물정보table_생성
    Call 담보물정보_범위변환(Array("담보물정보"))
    Call AdjustColumnWidth("2-1", 12, 20, 10.5, 17, 13, 8.5, 10, 7, 13, 13)
    Call Hyperlink_sheet("2-1", "Output_공시지가")
    Call Hyperlink_sheet("2-1", "Output_등본조회")
End Sub
Sub 실거래사례_1() ' 버튼명 : 실거래사례_국토
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("6-1")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call 실거래사례_국토("Output_거리_국토")
    Call 서식동기화("실거래사례_국토")
    Call 실거래사례_국토_범위변환(Array("실거래사례_국토"))
    Call AdjustColumnWidth("6-1", 5, 10, 40, 10, 12, 8.5, 17, 17)
    Call Hyperlink_sheet("6-1", "Output_거리_국토")
End Sub
Sub 실거래사례_2() ' 버튼명 : 실거래사례_밸류맵
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("6-2")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Exit Sub
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call 실거래사례_밸류맵("Output_거리_밸류맵")
    Call 서식동기화("실거래사례_밸류맵")
    Call 실거래사례_밸류맵_범위변환(Array("실거래사례_밸류맵"))
    Call AdjustColumnWidth("6-2", 5, 10, 40, 10, 15, 15, 11, 8.5, 17, 17)
    Call Hyperlink_sheet("6-2", "Output_거리_밸류맵")
End Sub
Sub 테이블범위변환() 'legacy code
    Call 담보물정보_실거래사례_병합(Array("담보물정보", "실거래사례_국토", "실거래사례_밸류맵"))
End Sub



