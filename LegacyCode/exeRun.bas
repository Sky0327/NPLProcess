Attribute VB_Name = "exeRun"
Function RunPythonScript(ScriptName As String) As Integer

    Dim exeFilePath As String
    Dim shellCommand As String
    Dim workbookPath As String

    On Error GoTo ErrorHandler ' 오류 발생 시 ErrorHandler로 이동

    ' 현재 매크로 파일이 저장된 경로 (상대경로로 설정)
    workbookPath = ThisWorkbook.Path
    exeFilePath = workbookPath & "\main.exe"
    
    ' 실행 파일이 발견되지 않으면 종료
    If Dir(exeFilePath) = "" Then
        MsgBox "실행파일을 찾을 수 없습니다. 코드 실행을 중단합니다."
        RunPythonScript = -1
        Exit Function
    End If

    ' CMD 창을 열고 main.exe와 scriptName을 실행하며 창이 닫히지 않도록 '/k' 옵션 사용
    ' 경로에 공백이 있을 경우를 대비하여 큰따옴표로 묶어줌
    shellCommand = exeFilePath & " " & ScriptName
    Debug.Print ("shellCommand: " & shellCommand)


    ' CMD 창을 띄우고 명령 실행, 창은 닫히지 않음
    Shell shellCommand, vbNormalFocus

    ' 정상적으로 실행되었음을 나타내는 값 반환
    RunPythonScript = 0
    Exit Function

ErrorHandler:
    MsgBox "코드가 중단되었습니다."
    RunPythonScript = -1
End Function

Function RunPythonScript_2(ScriptName As String) As Integer

    Dim exeFilePath As String
    Dim shellCommand As String
    Dim returnValue As Integer
    Dim workbookPath As String
    Dim wsh As Object ' WScript.Shell 객체 선언
    Dim shellExec As Object

    '현재 엑셀파일 저장
    ThisWorkbook.Save

    ' 현재 매크로 파일이 저장된 경로 (상대경로로 설정)
    workbookPath = ThisWorkbook.Path
    exeFilePath = workbookPath & "\main.exe"

    ' 파이썬 실행 파일이 발견되지 않으면 종료
    If Dir(exeFilePath) = "" Then
        MsgBox "실행파일을 찾을 수 없습니다. 코드 실행을 중단합니다."
        RunPythonScript_2 = -1
        Exit Function
    End If

    ' 파이썬 스크립트를 실행하기 전에 상태 표시줄에 메시지 표시
    Application.StatusBar = "★Smart_NPL_코드가 실행중입니다. 잠시만 기다려주세요.★"

    ' Shell을 사용하여 파이썬 스크립트를 실행
    shellCommand = exeFilePath & " " & ScriptName
    Set wsh = CreateObject("WScript.Shell")
    Set shellExec = wsh.Exec(shellCommand)
    
    ' 대기 시작 시간 기록
    waitStart = Timer

    ' 파이썬 코드가 종료될 때까지 대기 (0.2초마다 DoEvents를 호출하여 CPU 사용을 줄임)
    Do While shellExec.Status = 0
        If Timer - waitStart >= 0.2 Then
            DoEvents
            waitStart = Timer ' 대기 시작 시간 리셋
        End If
    Loop

    returnValue = shellExec.ExitCode ' 종료 코드 받기

    ' 상태 표시줄 메시지 복원
    Application.StatusBar = False

    ' 파이썬 코드가 종료된 후 결과 반환
    If returnValue = 0 Then
        MsgBox "코드 실행이 종료되었습니다. Output 불러오기를 실행해주세요.", vbInformation, "진행 상태"
    Else
        MsgBox "파이썬 코드 실행 중 오류가 발생했습니다. 오류 코드: " & returnValue
    End If

    RunPythonScript_2 = returnValue

End Function
Function RunPythonScript_3(ScriptName As String) As Integer
    Dim exeFilePath As String
    Dim shellCommand As String
    Dim returnValue As Integer
    Dim workbookPath As String
    Dim wsh As Object ' WScript.Shell 객체 선언

    '현재 엑셀파일 저장
    ThisWorkbook.Save

    ' 현재 매크로 파일이 저장된 경로 (상대경로로 설정)
    workbookPath = ThisWorkbook.Path
    exeFilePath = workbookPath & "\main.exe"

    ' 파이썬 실행 파일이 발견되지 않으면 종료
    If Dir(exeFilePath) = "" Then
        MsgBox "실행파일을 찾을 수 없습니다. 코드 실행을 중단합니다."
        RunPythonScript_3 = -1
        Exit Function
    End If

    ' 파이썬 스크립트를 실행하기 전에 상태 표시줄에 메시지 표시
    Application.StatusBar = "★Smart_NPL_코드가 실행중입니다. 잠시만 기다려주세요.★"

    ' Shell을 사용하여 파이썬 스크립트를 실행
    shellCommand = """" & exeFilePath & """ " & ScriptName
    Set wsh = CreateObject("WScript.Shell")
    returnValue = wsh.Run(shellCommand, 0, True) ' 세 번째 인수(True)는 프로세스가 종료될 때까지 대기함을 의미

    ' 상태 표시줄 메시지 복원
    Application.StatusBar = False

    ' 파이썬 코드가 종료된 후 결과 반환
    If returnValue = 0 Then
        MsgBox "코드 실행이 종료되었습니다. Output 불러오기를 실행해주세요.", vbInformation, "진행 상태"
    Else
        MsgBox "파이썬 코드 실행 중 오류가 발생했습니다. 오류 코드: " & returnValue
    End If

    RunPythonScript_3 = returnValue
End Function
' 단순 CMD 명령어 방식 + 최종산출물 감지 방식
Sub RunPythonScript_4(ScriptName As String, directoryName As String, filePrefix As String)
    Dim outputFilePath As String
    Dim isFileExists As Boolean
    Dim startTime As Double
    Dim timeout As Double
    Dim maxWaitTime As Double
    Dim sourceDir As String
    Dim fileName As String
    Dim pid As Double
    
    ' Source 시트에서 B4 셀 값 읽어오기
    sourceDir = ThisWorkbook.Sheets("Source").Range("B4").value
    
    ' 파일 경로는 B4 값과 인수로 받은 디렉토리명을 결합하여 생성
    outputFilePath = sourceDir & "\" & "Temp" & "\" & directoryName

    ' 기존 Output 파일 삭제 (기존 파일이 남아있으면 삭제)
    fileName = Dir(outputFilePath & "\" & filePrefix & "*.xlsx")
    Do While fileName <> ""
        Kill outputFilePath & "\" & fileName ' 기존 Output 파일 삭제
        fileName = Dir
    Loop

    ' 최대 대기 시간 설정 (예: 60초)
    maxWaitTime = 600 ' 600초 동안 대기
    
    ' 파이썬 스크립트를 비동기적으로 실행
    Call RunPythonScriptAsync(ScriptName)

    ' 현재 시간 기록
    startTime = Timer

    ' 파일이 생성될 때까지 대기 (filePrefix로 시작하는 파일이 있는지 탐색)
    Do
        DoEvents ' UI 업데이트 및 이벤트 처리
        
        ' 파일 탐색: 지정된 경로에서 filePrefix로 시작하는 파일이 있는지 확인
        fileName = Dir(outputFilePath & "\" & filePrefix & "*.xlsx")
        isFileExists = (fileName <> "")
        
        ' 일정 시간이 지나면 타임아웃 처리
        timeout = Timer - startTime
        If timeout > maxWaitTime Then
            MsgBox "파이썬 작업이 시간 내에 완료되지 않았습니다.", vbExclamation
            Exit Sub
        End If
    Loop Until isFileExists
    
    ' 파일이 생성되면 2초 대기 후 종료
    If isFileExists Then
        ' 상태 표시줄 메시지 복원
        Application.StatusBar = False
        Exit Sub
    End If
End Sub

Function RunPythonScriptAsync(ScriptName As String) As Integer
    Dim exeFilePath As String
    Dim shellCommand As String
    Dim pid As Double
    Dim workbookPath As String

    ' 현재 엑셀파일 저장
    ThisWorkbook.Save

    ' 절대 경로로 설정 (ThisWorkbook.Path를 사용해 경로 확인 후 절대 경로로 변경)
    exeFilePath = ThisWorkbook.Path & "\main.exe"

    ' 실행 파일 경로가 올바른지 확인
    If Dir(exeFilePath) = "" Then
        MsgBox "실행파일을 찾을 수 없습니다. 경로: " & exeFilePath
        RunPythonScriptAsync = -1
        Exit Function
    End If

    ' 파이썬 스크립트를 실행하기 전에 상태 표시줄에 메시지 표시
    Application.StatusBar = "★Smart_NPL_코드가 실행중입니다. 파이썬 코드가 백그라운드에서 실행됩니다.★"

    ' Shell을 사용하여 파이썬 스크립트를 비동기 실행
    shellCommand = """" & exeFilePath & """ " & ScriptName
    pid = Shell(shellCommand, vbNormalFocus) ' 비동기 실행

    ' 정상적으로 프로세스가 시작된 후 바로 반환
    If pid <> 0 Then
    Else
        MsgBox "파이썬 코드 실행 중 문제가 발생했습니다.", vbExclamation, "에러"
    End If

    ' 반환값 설정
    RunPythonScriptAsync = 0
End Function

Sub run_kb_info()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_KB시세")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_kb_info", "KB시세", "Output_")
End Sub

Sub run_court_auction()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_법원경매")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_court_auction", "법원경매", "Output_")
End Sub

Sub run_infocare_analysis()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_인포통계")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_infocare_analysis", "인포통계", "Output_")
End Sub

Sub run_infocare_integrated()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_인포통합")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_infocare_integrated", "인포통합", "Output_")
End Sub
Sub run_infocare_case()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_인포사례상세")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_infocare_case", "인포사례상세", "Output_")
End Sub
