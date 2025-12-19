Attribute VB_Name = "Call_Main"
Option Explicit
Sub main_initialize_file(control As IRibbonControl)
    Call InitializeFile
End Sub
Sub main_set_report_name(control As IRibbonControl)
    Call OpenfrmSetReport
End Sub
Sub main_set_id(control As IRibbonControl)
    Call OpenfrmSetID
End Sub
Sub main_get_registry_info_basic_삭제() '더이상 사용하지 않음
    Call OpenfrmRegistryInfo '
End Sub
Sub main_get_registry_info_basic(control As IRibbonControl)
    Call Run_run_get_registry_basic_info
    Call Output_불러오기_등본목록
End Sub

Sub main_set_input_1(control As IRibbonControl) 'Input 세팅 - Input 불러오기
    Call makeInputSheet_input
    Call set_input_main_1
    Call MatchAndExtract
    Call LookupAndCopy
    Call CopyFormulaToLastRow
End Sub
Sub main_set_input_2(control As IRibbonControl)
    Call SpeedUp
    Call makeInputSheet_등본조회
    Call set_input_2
    Call SpeedDown
End Sub
Sub main_set_input_2_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_등본조회
End Sub
Sub main_get_register_hyphen(control As IRibbonControl)
    Call CheckRequiredSheet("Input_등본조회")
    Call Run_run_register_inquiry
    Call Output_불러오기_등본조회
End Sub
Sub main_set_input_3(control As IRibbonControl)
    Call SpeedUp
    Call makeInputSheet_공시지가
    Call set_input_3
    Call SpeedDown
End Sub
Sub main_set_input_3_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_공시지가
End Sub
Sub main_posting_price(control As IRibbonControl)
    Call CheckRequiredSheet("Input_공시지가")
    Call Run_run_posting_price
    Call Output_불러오기_공시지가
End Sub
Sub main_set_input_4(control As IRibbonControl)
    Call SpeedUp
    Call makeInputSheet_실거래가
    Call set_input_4
    'If ThisWorkbook.Sheets("Input_실거래가").Range("B7").value = "" Then
    '    Exit Sub
    'End If
    Call MatchAndExtract2
    Call MatchAndExtract3
    Call MarkV
    'Call MarkUniqueValuesByGroup
    Call SetValuesBasedOnKColumn
    Call DoubleLookupAndReturnValuesWithLoops
    Call SpeedDown
End Sub
Sub main_set_input_4_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_실거래가
End Sub
Sub main_level_update(control As IRibbonControl)
    Call CheckRequiredSheet("Input_실거래가")
    Call DoubleLookupAndReturnValuesWithLoops
    MsgBox "(밸류맵)입력된 확장레벨에 따른 인근 동 범위가 업데이트되었습니다."
End Sub
Sub main_gookto_real_price(control As IRibbonControl)
    Call CheckRequiredSheet("Input_실거래가")
    Call Run_run_gookto_real_price
    Call Output_불러오기_실거래가조회_국토교통부
End Sub
Sub main_valuemap(control As IRibbonControl)
    Call CheckRequiredSheet("Input_실거래가")
    Call Run_run_valuemap
    Call Output_불러오기_실거래가조회_밸류맵
End Sub
Sub main_gookto_distance(control As IRibbonControl)
    '선행시트 검증 이미 아래 run에 있으므로 생략
    Call Run_run_gookto_calculate_distance
    Call Output_불러오기_거리계산_국토교통부
End Sub
Sub main_valuemap_distance(control As IRibbonControl)
    '선행시트 검증 이미 아래 run에 있으므로 생략
    Call Run_run_valuemap_calculate_distance
    Call Output_불러오기_거리계산_밸류맵
End Sub
Sub main_set_input_5(control As IRibbonControl)
    Call CheckRequiredSheet("Input")
    Call writeInputKB
End Sub
Sub main_set_input_5_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_KB시세
End Sub
Sub main_KB(control As IRibbonControl)
    Call CheckRequiredSheet("Input_KB시세")
    Call run_kb_info
    Call Output_불러오기_KB시세
End Sub
Sub main_set_input_6(control As IRibbonControl)
    Call writeInputCourt
End Sub
Sub main_set_input_6_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_법원경매
End Sub
Sub main_court(control As IRibbonControl)
    Call CheckRequiredSheet("Input_법원경매")
    Call run_court_auction
    Call Output_불러오기_법원경매
End Sub
Sub main_set_input7(control As IRibbonControl) 'Input 세팅(통계)
    Call writeInputInfoAnalysis
End Sub
Sub main_set_input7_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_인포통계
End Sub
Sub main_set_input8(control As IRibbonControl) 'Input 세팅(통합)
    Call writeInputInfoAll
End Sub
Sub main_set_input8_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_인포통합
End Sub
Sub main_set_input9(control As IRibbonControl) 'Input 세팅(사례)
    Call makeInputSheetInfoDetail
End Sub
Sub main_set_input9_Empty(control As IRibbonControl)
    Call CreateEmptyInputSheet_인포사례
End Sub
Sub main_infocare_analysis(control As IRibbonControl) '조회(통계)
    Call CheckRequiredSheet("Input_인포통계")
    
    ' Input_인포통계 시트의 G열에서 "전체" 텍스트 검색
    If WorksheetFunction.CountIf(Sheets("Input_인포통계").columns("G"), "전체") > 0 Then
        MsgBox "인포케어 통계 조회는 '전체' 지역 검색을 제공하지 않습니다.", vbCritical
        Call GoEnd
    End If

    Call run_infocare_analysis
    Call Output_불러오기_인포통계
End Sub
Sub main_infocare_integrated(control As IRibbonControl) '조회(통합)
    Call CheckRequiredSheet("Input_인포통합")
    Call run_infocare_integrated
    Call Output_불러오기_인포통합
End Sub
Sub main_infocare_case(control As IRibbonControl) '조회(사례)
    Call CheckRequiredSheet("Input_인포사례상세")
    Call run_infocare_case
    Call Output_불러오기_인포사례상세
End Sub










