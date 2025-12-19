Attribute VB_Name = "Call_HJ"
Option Explicit
'유저폼 오픈
Sub OpenfrmRegistryInfo()
    FrmRegistryInfo.Show
End Sub
Sub OpenfrmSetReport()
    FrmSetReport.Show
End Sub
Sub OpenfrmSetID()
    FrmSetID.Show
End Sub
'Input sheet 생성
Sub makeInputSheet_input()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input", "Input")
End Sub
Sub makeInputSheet_등본조회()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_등본조회")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_등본조회", "Input_등본조회")
End Sub
Sub makeInputSheet_공시지가()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_공시지가")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_공시지가", "Input_공시지가")
End Sub
Sub makeInputSheet_실거래가()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_실거래가")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_실거래가", "Input_실거래가")
End Sub
Sub makeInputSheet_KB시세()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_법원경매")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_KB시세", "Input_KB시세")
End Sub
Sub makeInputSheet_법원경매()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_법원경매")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_법원경매", "Input_법원경매")
End Sub
Sub makeInputSheet_인포통합()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_인포통합")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_인포통합", "Input_인포통합")
End Sub
Sub makeInputSheet_인포통계()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_인포통계")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_인포통계", "Input_인포통계")
End Sub
Sub makeInputSheet_인포사례상세()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Input_인포사례상세")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call makeInputSheet("Tpl_Input_인포사례상세", "Input_인포사례상세")
End Sub
'파이썬 스크립트 실행
Sub Run_run_get_registry_basic_info()
    Dim SheetExists As Boolean
    Dim FolderPicker As FileDialog
    Dim folderPath As String

    ' 폴더 선택 창 띄우기
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FolderPicker
        .Title = "등기부등본PDF가 포함된 폴더를 선택하세요. 모든 하위폴더를 탐색합니다."
        .AllowMultiSelect = False ' 다중 선택 불가
        If .Show = -1 Then ' 사용자가 폴더를 선택했을 경우
            folderPath = .SelectedItems(1) ' 선택된 폴더 경로 저장
            shtSource.Cells(shtSource.columns(1).Find("등기부등본_input폴더경로").row, 2).value = folderPath
        Else
            Call GoEnd
        End If
    End With

    SheetExists = Check_Sheet("Output_등본목록")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    
    ThisWorkbook.Save
    
    Call RunPythonScript_4("run_get_registry_basic_info", "등기부등본 기본정보", "Output_")
End Sub
Sub Run_run_valuemap()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_실거래가_밸류맵")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Call RunPythonScript_4("run_valuemap", "실거래가조회_밸류맵", "Output_")
End Sub
Sub Run_run_register_inquiry()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_등본조회")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Dim SheetExists_2 As Boolean
    SheetExists_2 = Check_Sheet("Output_등본조회(전체)")
    If SheetExists_2 = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    If ValidateTable("Input_등본조회", "C", "D") Then
        Call RunPythonScript_4("run_register_inquiry", "등기부등본", "Output_")
    Else
        Call GoEnd
    End If
End Sub
Sub Run_run_posting_price()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_공시지가")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    Dim SheetExists_3 As Boolean
    SheetExists_3 = Check_Sheet("Output_공시지가(전체)")
    If SheetExists_3 = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    If ValidateTable("Input_공시지가", "C", "D") Then
        Call RunPythonScript_4("run_posting_price", "공시지가", "Output_")
    Else
        Call GoEnd
    End If
End Sub
Sub Run_run_gookto_real_price()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_실거래가_국토")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    If ValidateTable("Input_실거래가", "C", "D", "E", "F", "G", "H", "I", "J", "K") Then
        Call RunPythonScript_4("run_gookto_real_price", "실거래가조회_국토교통부", "Output_")
    Else
        Call GoEnd
    End If
End Sub
Sub Run_run_gookto_calculate_distance()
    Dim SheetExists As Boolean
    SheetExists = Check_Sheet("Output_거리_국토")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    If ValidateTable("Output_실거래가_국토", "E", "F", "J", "K", "L", "M") Then
        Call RunPythonScript_4("run_gookto_calculate_distance", "거리계산_국토교통부", "Output_")
    Else
        Call GoEnd
    End If
End Sub
Sub Run_run_valuemap_calculate_distance()
    Dim SheetExists As Boolean

    SheetExists = Check_Sheet("Output_거리_밸류맵")
    If SheetExists = False Then
        '유저가 "아니오"를 누른 경우 전체 Sub 종료
        Call GoEnd
    End If '유저가 "예"를 누른 경우 기존 시트 삭제 후 후속 코드 진행
    If ValidateTable("Output_실거래가_밸류맵", "B", "C") Then
        Call RunPythonScript_4("run_valuemap_calculate_distance", "거리계산_밸류맵", "Output_")
    Else
        Call GoEnd
    End If
End Sub
