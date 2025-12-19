Attribute VB_Name = "Module_KeyTest"
Option Explicit
' TestKeyTest_hyphen1 함수 설명:
' KeyTest_hyphen1(uniq_no, hyphen_id, hyphen_apikey)
'    uniq_no      : 부동산 고유번호 (예: "2241-1996-703245")
'    hyphen_id    : Source 시트 A열에서 "hyphen_id"가 있는 행의 오른쪽(Offset(0,1)) 셀 값
'    hyphen_apikey: Source 시트 A열에서 "hyphen_apikey"가 있는 행의 오른쪽(Offset(0,1)) 셀 값
'
' API 호출 결과는 KeyTest 시트에서 A열에 "KeyTest_Hyphen1"가 있는 행의 오른쪽 오른쪽(Offset(0,2)) 셀에 반영됩니다.
Sub TestKey_hyphen1()
    Dim wsSource As Worksheet      ' Source 시트 참조
    Dim wsKeyTest As Worksheet     ' KeyTest 시트 참조
    Dim hyphen_id As String        ' hyphen_id 값을 저장할 변수
    Dim hyphen_apikey As String    ' hyphen_apikey 값을 저장할 변수
    Dim uniq_no As String          ' 부동산 고유번호 (예: "2241-1996-703245")
    Dim result As String           ' API 호출 결과를 저장할 변수
    
    ' 현재 Workbook의 Source 및 KeyTest 시트를 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    hyphen_id = wsSource.columns("A").Find("hyphen_id", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    hyphen_apikey = wsSource.columns("A").Find("hyphen_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    uniq_no = "2241-1996-703245"
    
    result = KeyTest_hyphen1(uniq_no, hyphen_id, hyphen_apikey)
    
    wsKeyTest.columns("A").Find("KeyTest_Hyphen1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_Hyphen1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' API 호출 함수
Public Function KeyTest_hyphen1(uniq_no As String, _
                                hyphen_id As String, _
                                hyphen_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim jsonBody As String
    Dim statusCode As Long
    Dim responseText As String
    
    ' API 요청 URL
    url = "https://api.hyphen.im/in0004000169"
    
    ' JSON 형식의 요청 바디
    jsonBody = "{""uniqNo"": """ & uniq_no & """}"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

    ' POST 방식의 동기 요청
    xmlhttp.Open "POST", url, False
    
    ' 헤더 설정
    xmlhttp.setRequestHeader "hyphen-gustation", "Y"
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "user-id", hyphen_id
    xmlhttp.setRequestHeader "Hkey", hyphen_apikey
    
    Debug.Print "등기부등본 API를 호출합니다... (고유번호: " & uniq_no & ")"
    xmlhttp.send jsonBody
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부를 판단하고 전체 content를 반환
    If statusCode = 200 Then
        KeyTest_hyphen1 = "호출결과: " & responseText
    Else
        KeyTest_hyphen1 = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "API 요청 중 예외 발생 (고유번호: " & uniq_no & "): " & Err.Description
    KeyTest_hyphen1 = "호출실패: " & Err.Description
End Function




' Kakao API를 호출하는 테스트 Sub
Sub TestKey_KakaoAPI()
    Dim wsSource As Worksheet      ' Source 시트 참조
    Dim wsKeyTest As Worksheet     ' KeyTest 시트 참조
    Dim kakao_apikey As String     ' Kakao API Key 값
    Dim address As String          ' 조회할 주소
    Dim result As String           ' API 호출 결과
    
    ' 현재 Workbook의 Source 및 KeyTest 시트를 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' Kakao API Key 가져오기
    kakao_apikey = wsSource.columns("A").Find("kakao_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    ' 조회할 주소 입력 (테스트 주소)
    address = "서울특별시 용산구 한강대로 100"
    
    ' Kakao API 호출
    result = KeyTest_Kakao(address, kakao_apikey)
    
    ' API 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_KakaoAPI", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_KakaoAPI", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' Kakao 주소 검색 API 호출 함수
Public Function KeyTest_Kakao(address As String, kakao_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim responseText As String
    Dim statusCode As Long
    
    ' Kakao API 요청 URL 설정
    url = "https://dapi.kakao.com/v2/local/search/address.json?query=" & URLEncode(address)
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식으로 API 요청
    xmlhttp.Open "GET", url, False
    ' 헤더 설정
    xmlhttp.setRequestHeader "Authorization", "KakaoAK " & kakao_apikey
    
    Debug.Print "Kakao API를 호출합니다... (주소: " & address & ")"
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드 확인
    If statusCode = 200 Then
        KeyTest_Kakao = "호출결과: " & responseText
    Else
        KeyTest_Kakao = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "API 요청 중 예외 발생 (주소: " & address & "): " & Err.Description
    KeyTest_Kakao = "호출실패: " & Err.Description
End Function

' URL 인코딩 함수 (한글 주소 처리)
Function URLEncode(str As String) As String
    Dim i As Integer
    Dim ascValue As String
    Dim encodedStr As String
    
    encodedStr = ""
    
    For i = 1 To Len(str)
        ascValue = Mid(str, i, 1)
        Select Case Asc(ascValue)
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9, A-Z, a-z
                encodedStr = encodedStr & ascValue
            Case Else
                encodedStr = encodedStr & "%" & Right("0" & Hex(Asc(ascValue)), 2)
        End Select
    Next i
    
    URLEncode = encodedStr
End Function





' 공동주택 API 호출 함수
Public Function CallTownAPI(unique_id As String, _
                            dong_number As String, _
                            hosu_number As String, _
                            twin_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim params As String
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "http://api.vworld.kr/ned/data/getApartHousingPriceAttr"
    
    ' GET 파라미터 설정
    params = "key=" & twin_apikey & _
             "&pnu=" & unique_id & _
             "&stdrYear=2024" & _
             "&dongNm=" & dong_number & _
             "&hoNm=" & hosu_number & _
             "&format=json" & _
             "&numOfRows=1" & _
             "&pageNo=1"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url & "?" & params, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        CallTownAPI = "호출결과: " & responseText
    Else
        CallTownAPI = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    CallTownAPI = "호출실패: " & Err.Description
End Function

' 단독주택 API 호출 함수
Public Function CallIndividualAPI(unique_id As String, twin_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim params As String
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "https://api.vworld.kr/ned/data/getIndvdHousingPriceAttr"
    
    ' GET 파라미터 설정
    params = "key=" & twin_apikey & _
             "&pnu=" & unique_id & _
             "&stdrYear=2024" & _
             "&format=json" & _
             "&numOfRows=1" & _
             "&pageNo=1"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url & "?" & params, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        CallIndividualAPI = "호출결과: " & responseText
    Else
        CallIndividualAPI = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    CallIndividualAPI = "호출실패: " & Err.Description
End Function

' 토지 API 호출 함수
Public Function CallLandAPI(unique_id As String, twin_apikey As String, year As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim params As String
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "https://api.vworld.kr/ned/data/getLandCharacteristics"
    
    ' GET 파라미터 설정 (연도별 데이터 요청)
    params = "key=" & twin_apikey & _
             "&pnu=" & unique_id & _
             "&stdrYear=" & year & _
             "&format=json" & _
             "&numOfRows=1" & _
             "&pageNo=1"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url & "?" & params, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        CallLandAPI = "호출결과: " & responseText
    Else
        CallLandAPI = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    CallLandAPI = "호출실패: " & Err.Description
End Function





' 공동주택 API 호출을 테스트하는 Sub
Sub TestKey_Twin1()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim twin_apikey As String
    Dim unique_id As String
    Dim dong_number As String
    Dim hosu_number As String
    Dim result As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    twin_apikey = wsSource.columns("A").Find("twin_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    unique_id = "1234567890"  ' 테스트용 부동산 고유번호
    dong_number = "101"
    hosu_number = "202"

    ' API 호출
    result = CallTownAPI(unique_id, dong_number, hosu_number, twin_apikey)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' 단독주택 API 호출을 테스트하는 Sub
Sub TestKey_Twin2()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim twin_apikey As String
    Dim unique_id As String
    Dim result As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    twin_apikey = wsSource.columns("A").Find("twin_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    unique_id = "1234567890"  ' 테스트용 부동산 고유번호

    ' API 호출
    result = CallIndividualAPI(unique_id, twin_apikey)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI2", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI2", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' 토지 API 호출을 테스트하는 Sub
Sub TestKey_Twin3()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim twin_apikey As String
    Dim unique_id As String
    Dim result As String
    Dim year As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    twin_apikey = wsSource.columns("A").Find("twin_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    unique_id = "1234567890"  ' 테스트용 부동산 고유번호
    year = "2024"  ' 연도 선택

    ' API 호출
    result = CallLandAPI(unique_id, twin_apikey, year)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI3", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_TwinAPI3", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub









' 아파트 실거래가 API 테스트 함수
Sub TestKey_data1()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim data_apikey As String
    Dim gu_code As String
    Dim base_date As String
    Dim result As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    data_apikey = wsSource.columns("A").Find("data_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    gu_code = "11110"  ' 테스트용 구 코드
    base_date = "202401"  ' 테스트용 기준 날짜

    ' API 호출
    result = GetDataApt(gu_code, base_date, data_apikey)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_data_apikey1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_data_apikey1", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' 오피스텔 실거래가 API 테스트 함수
Sub TestKey_data2()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim data_apikey As String
    Dim gu_code As String
    Dim base_date As String
    Dim result As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    data_apikey = wsSource.columns("A").Find("data_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    gu_code = "11110"  ' 테스트용 구 코드
    base_date = "202401"  ' 테스트용 기준 날짜

    ' API 호출
    result = GetDataOffice(gu_code, base_date, data_apikey)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_data_apikey2", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_data_apikey2", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub

' 연립/다세대 실거래가 API 테스트 함수
Sub TestKey_data3()
    Dim wsSource As Worksheet
    Dim wsKeyTest As Worksheet
    Dim data_apikey As String
    Dim gu_code As String
    Dim base_date As String
    Dim result As String
    
    ' Source 및 KeyTest 시트 지정
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsKeyTest = ThisWorkbook.Sheets("KeyTest")
    
    ' API Key 및 파라미터 설정
    data_apikey = wsSource.columns("A").Find("data_apikey", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 1).value
    gu_code = "11110"  ' 테스트용 구 코드
    base_date = "202401"  ' 테스트용 기준 날짜

    ' API 호출
    result = GetDataMulti(gu_code, base_date, data_apikey)

    ' 결과 저장
    wsKeyTest.columns("A").Find("KeyTest_data_apikey3", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 2).value = result
    wsKeyTest.columns("A").Find("KeyTest_data_apikey3", LookIn:=xlValues, LookAt:=xlWhole).offset(0, 3).value = Now
End Sub


' 아파트 실거래가 API 호출 함수
Public Function GetDataApt(gu_code As String, base_date As String, data_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade" & _
          "?LAWD_CD=" & gu_code & "&DEAL_YMD=" & base_date & "&serviceKey=" & data_apikey & _
          "&pageNo=1&numOfRows=10000"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        GetDataApt = "호출결과: " & responseText
    Else
        GetDataApt = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    GetDataApt = "호출실패: " & Err.Description
End Function

' 오피스텔 실거래가 API 호출 함수
Public Function GetDataOffice(gu_code As String, base_date As String, data_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "https://apis.data.go.kr/1613000/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade" & _
          "?serviceKey=" & data_apikey & "&LAWD_CD=" & gu_code & "&DEAL_YMD=" & base_date & _
          "&pageNo=1&numOfRows=10000"
    
    On Error GoTo ErrHandler
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        GetDataOffice = "호출결과: " & responseText
    Else
        GetDataOffice = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    GetDataOffice = "호출실패: " & Err.Description
End Function

' 연립/다세대 실거래가 API 호출 함수
Public Function GetDataMulti(gu_code As String, base_date As String, data_apikey As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim responseText As String
    Dim statusCode As Long
    
    ' API 요청 URL
    url = "https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade" & _
          "?serviceKey=" & data_apikey & "&LAWD_CD=" & gu_code & "&DEAL_YMD=" & base_date & _
          "&pageNo=1&numOfRows=10000"
    Debug.Print "url :" & url
    
    On Error GoTo ErrHandler
    '시도1
    'Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    '시도2
    'Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    '시도3
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' GET 방식의 동기 요청
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    statusCode = xmlhttp.Status
    responseText = xmlhttp.responseText
    'Debug.Print responseText
    
    ' HTTP 상태 코드에 따라 성공 여부 판단
    If statusCode = 200 Then
        GetDataMulti = "호출결과: " & responseText
    Else
        GetDataMulti = "호출실패: HTTP 상태 코드 " & statusCode & " - " & responseText
    End If
    
    Exit Function

ErrHandler:
    GetDataMulti = "호출실패: " & Err.Description
End Function


