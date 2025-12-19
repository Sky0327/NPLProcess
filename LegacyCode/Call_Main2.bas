Attribute VB_Name = "Call_Main2"
Option Explicit

Sub main_report_0(control As IRibbonControl)
    Call CheckRequiredSheet("Input")
    Call makeReportSheet_0
    Call removeApo
    Call MoveSheetToLast("0")
End Sub

Sub main_report_1(control As IRibbonControl)
    Call makeReportSheet_1
    Call ImportSellList
    Call MoveSheetToLast("1")
End Sub
Sub main_report_mapping(control As IRibbonControl) ' 버튼명 [0]담보 Mapping
    Call CheckRequiredSheet("Input")
    Call 담보물정보_생성
End Sub

Sub main_report_2_1(control As IRibbonControl) ' 버튼명 [2-1]담보물정보 생성
    Call CheckRequiredSheet("Input")
    Call CheckRequiredSheet("Output_등본조회")
    Call CheckRequiredSheet("Output_공시지가")
    If Not CheckCollateralMapping() Then
        Exit Sub ' Mapping이 완전하지않으면 Sub 종료
    End If
    Call 담보물정보_table_생성
    Call MoveSheetToLast("2-1")
End Sub

Sub main_report_2_2(control As IRibbonControl)
    Call CheckRequiredSheet("Output_KB시세")
    Call UpdateTableKB
    Call AdjustColumnWidth("2-2", 3, 12, 8.5, 8.5, 15, 15)
    Call MoveSheetToLast("2-2")
End Sub

Sub main_report_3(control As IRibbonControl)
    Call CheckRequiredSheet("Output_법원경매")
    Call UpdateTableAuction
    Call AdjustColumnWidth("3", 3, 12, 13, 10, 16, 12, 15)
    Call MoveSheetToLast("3")
End Sub

Sub main_report_4(control As IRibbonControl)
    MsgBox ("현재는 지원되지 않는 기능입니다.")
End Sub

Sub main_report_5_1(control As IRibbonControl)
    Call CheckRequiredSheet("Output_인포통계")
    Call UpdateAnalysisTable
    Call AdjustColumnWidth("5-1", 3)
    Call MoveSheetToLast("5-1")
End Sub

Sub main_report_5_2(control As IRibbonControl)
    Call CheckRequiredSheet("Output_인포사례상세")
    Call UpdateTableCases
    Call MoveSheetToLast("5-2")
End Sub
Sub main_report_6_1(control As IRibbonControl) ' 버튼명 [6-1]거래사례 국토
    If Not CheckForVInColumn("Output_거리_국토", "U") Then
        Exit Sub ' V가 없으면 Sub 종료
    End If
    Call CheckRequiredSheet("Output_거리_국토")
    Call 실거래사례_1
    Call MoveSheetToLast("6-1")
End Sub
Sub main_report_6_2(control As IRibbonControl) ' 버튼명 [6-2]거래사례 밸류맵
    If Not CheckForVInColumn("Output_거리_밸류맵", "O") Then
        Exit Sub ' V가 없으면 Sub 종료
    End If
    Call CheckRequiredSheet("Output_거리_밸류맵")
    Call 실거래사례_2
    Call MoveSheetToLast("6-2")
End Sub

Sub main_ConvertToXlsx(control As IRibbonControl)
    Call ConvertToXlsx
End Sub
Sub main_ConvertToXlsx2(control As IRibbonControl)
    Call ConvertToXlsx
End Sub

Sub main_create_input_changePDFName(control As IRibbonControl)
    Call CheckRequiredSheet("Output_등본목록")
    ThisWorkbook.Sheets("Output_등본목록").Activate
    Call GetToBePDFName
    Call AdjustColumnWidth("Output_등본목록", 8, 70, 17, 15, 70, 70)
    Call ApplyColorFormatting_grey("Output_등본목록", "A", "B")
    Call ApplyColorFormatting_vividpersimon("Output_등본목록", "C", "D", "E")
    Call ChangeFontSize("Output_등본목록")
    Call ApplyAllBorders("Output_등본목록")
End Sub

Sub main_execute_conversion(control As IRibbonControl)
    Call CheckRequiredSheet("Output_등본목록")
    Call RenamePDFFiles
    
End Sub

