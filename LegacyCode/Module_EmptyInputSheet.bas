Attribute VB_Name = "Module_EmptyInputSheet"
Option Explicit
Sub CreateEmptyInputSheet_등본조회()
    If Check_Sheet("Input_등본조회") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_등본조회", "Input_등본조회")
End Sub
Sub CreateEmptyInputSheet_공시지가()
    If Check_Sheet("Input_공시지가") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_공시지가", "Input_공시지가")
End Sub
Sub CreateEmptyInputSheet_실거래가()
    If Check_Sheet("Input_실거래가") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_실거래가", "Input_실거래가")
End Sub
Sub CreateEmptyInputSheet_KB시세()
    If Check_Sheet("Input_KB시세") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_KB시세", "Input_KB시세")
End Sub
Sub CreateEmptyInputSheet_법원경매()
    If Check_Sheet("Input_법원경매") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_법원경매", "Input_법원경매")
    ThisWorkbook.Sheets("Input_법원경매").columns("E").Delete
End Sub
Sub CreateEmptyInputSheet_인포통계()
    If Check_Sheet("Input_인포통계") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_인포통계", "Input_인포통계")
    ThisWorkbook.Sheets("Input_인포통계").columns("H").Delete
End Sub
Sub CreateEmptyInputSheet_인포통합()
    On Error Resume Next
    If Check_Sheet("Input_인포통합") = False Then: GoEnd
    Call makeInputSheet("Tpl_Input_인포통합", "Input_인포통합")
    ThisWorkbook.Sheets("Input_인포통합").columns("H").Delete
End Sub
Sub CreateEmptyInputSheet_인포사례()
    Msg "Input 세팅(사례) 버튼을 이용 부탁드립니다."
    Call GoEnd
    'Call makeInputSheetInfoDetail
    'If Check_Sheet("Input") = False Then: GoEnd
    'Call makeInputSheet("Tpl_Input", "Input_인포사례상세")
End Sub
