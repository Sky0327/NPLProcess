Attribute VB_Name = "PwC"
Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwLen As Long, ByRef pbBuffer As Byte) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Public Const SRVer As String = "20250207" '버전 여기에 업데이트
Public Const RelVersion As String = "R"
Public Const ThisFontName  As String = "맑은 고딕"
Public Const ThisFontSize  As Integer = 9
Public Const ThisRowHeight  As Integer = 18
Public Const ThisColumnWidth As Integer = 16

Type udtRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As udtRECT) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As udtRECT) As Long
#End If

Function IsPwC() As Boolean
    On Error Resume Next
    If InStr(UCase(Environ("Userdnsdomain")), "PWC") > 0 Then IsPwC = True
End Function

Sub TestIsPwC2()
    If IsPwC() Then
        Msg "PwC 도메인에 속해 있습니다."
        Msg Environ("Userdnsdomain")
    Else
        Msg "PwC 도메인에 속해 있지 않습니다."
    End If
End Sub

Sub WBOpen()
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim StrMsg As String
    If IsPwC Then
    LastPwC.value = GetUserMSMail
    TimeToGoogle True
    'Msg "PwC 계정으로 접속되었습니다."
    Else
        TimeToGoogle True
        Msg "미인가 사용자입니다." & vbLf & "PwC 인증 컴퓨터 외에서는 사용이 불가능합니다."
        ThisWorkbook.Close: Exit Sub
    End If
    
    Call AddOpenTime
    
    'Call SpeedUp: Call CloseWindow: Call SpeedDown: Call AddOpenTime
    'If Now - RelDate > 180 Then Msg "해당 버전은 유효기간이 6개월 이상 경과되었습니다." & vbLf & "최신 스마트리뷰어를 다운 받아 작업하시기를 권장드립니다."
    
End Sub

Public Function GetUserMSMail() As String
    On Error Resume Next
    GetUserMSMail = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity\ADUserName")
End Function

Public Function LastPwC() As Range
    Set LastPwC = Sources.Cells(11, 1)
End Function

Sub test2()

TimeToGoogle True

End Sub

Sub TimeToGoogle(Optional ByVal IsMandatory As Boolean, Optional ByVal UpdateTime As Boolean, Optional ByVal StrFilePath As String)
    On Error Resume Next
    
    Dim FormURL1 As String, FormURL2 As String, FormURL3 As String
    Dim StrForm As String, StrEncForm As String, OpenValue As Double, CloseValue As Double, UsageValue As Double
    Dim UsageTime As String, StrOpen As String, StrClose As String, StrUsage As String, StrDate As String
    Dim StrEncOpen As String, StrEncClose As String, StrEncUsage As String, StrEncDate As String
    Dim StrEncName As String, StrEncMail As String, StrEncCompanyName As String, StrEncFilePath As String, StrEncPeriod As String, StrEncPwC As String, StrEncVer As String
    Dim StrName As String, StrMail As String, StrVer As String, StrCompanyName As String, StrPeriod As String, StrPwC As String

    ' 구글 폼 URL
    FormURL1 = "https://docs.google.com/forms/d/e/1FAIpQLSeV5hx_N_NDOTiCrWjy7hJM1gUUl5_GWkWmMz-7IdnQjqO3Gw/formResponse?ifq"
    FormURL2 = "https://docs.google.com/forms/d/1-mGcERx-3hdSbbYvgodOGImD6yPRNw4SzmqjkJ6Fp7Y/formResponse?ifq"
    'FormURL3 = "https://docs.google.com/forms/d/e/1FAIpQLSfv861lafhzdACMixpexBlHELNwdvDvqu0YJ0HmESkMVS45sg/formResponse?ifq"
    
    ' 시간 값 계산
    CloseValue = Now
    OpenValue = GetOpenTime
    If OpenValue = 0 Then OpenValue = Sources.Cells(5, 3).value
    UsageValue = CloseValue - OpenValue
    If IsMandatory = False And UsageValue < 0.002082 Then Exit Sub
    
    ' 데이터 매핑 (새로운 Entry ID 반영)
    StrDate = "&entry.258382440=" & Format(CloseValue, "yyyy-MM-dd")
    StrEncDate = "&entry.258382440=" & Encrypt_form(Format(CloseValue, "yyyy-MM-dd"))
    
    StrName = "&entry.1771454798=" & ENCODEURL(UserName)
    StrEncName = "&entry.1771454798=" & Encrypt_form(ENCODEURL(UserName))
    
    StrMail = "&entry.1809925046=" & UserMail
    StrEncMail = "&entry.1809925046=" & Encrypt_form(UserMail)
    
    StrVer = "&entry.318203784=" & StrSRVer & "(" & Format(RelDate, "yyyy-mm-dd") & ")"
    StrEncVer = "&entry.318203784=" & Encrypt_form(StrSRVer & "(" & Format(RelDate, "yyyy-mm-dd") & ")")
    
    If CompanyName.value = vbNullString Then
        StrCompanyName = ""
    Else
        StrCompanyName = "&entry.11648735=" & URLAbbName
        StrEncCompanyName = "&entry.11648735=" & Encrypt_form(URLAbbName)
    End If
    
    StrPeriod = "&entry.560174368=" & "period.Value"
    StrEncPeriod = "&entry.560174368=" & Encrypt_form("period.Value")
    
    StrOpen = "&entry.2101217174=" & Format(OpenValue, "yyyy-MM-dd hh:mm:ss")
    StrEncOpen = "&entry.2101217174=" & Encrypt_form(Format(OpenValue, "yyyy-MM-dd hh:mm:ss"))
    
    StrClose = "&entry.634670804=" & Format(CloseValue, "yyyy-MM-dd hh:mm:ss")
    StrEncClose = "&entry.634670804=" & Encrypt_form(Format(CloseValue, "yyyy-MM-dd hh:mm:ss"))
    
    StrUsage = "&entry.486042821=" & Format(UsageValue, "hh:mm:ss")
    StrEncUsage = "&entry.486042821=" & Encrypt_form(Format(UsageValue, "hh:mm:ss"))
    
    StrPwC = "&entry.296399675=" & LastPwC.value
    StrEncPwC = "&entry.296399675=" & Encrypt_form(LastPwC.value)
    
    If StrFilePath = vbNullString Then
        StrFilePath = "&entry.2100198753=" & URLFilePathName
    Else
        StrFilePath = "&entry.2100198753=" & StrFilePath
    End If
    
    StrEncFilePath = "&entry.2100198753=" & Encrypt_form(StrFilePath)

    ' 데이터 제출용 폼 구성
    StrForm = StrDate & StrName & StrMail & StrVer & StrCompanyName & StrPeriod & StrOpen & StrClose & StrUsage & StrFilePath & StrPwC & "&submit=submit"
    StrEncForm = StrEncDate & StrEncName & StrEncMail & StrEncVer & StrEncCompanyName & StrEncPeriod & StrEncOpen & StrEncClose & StrEncUsage & StrEncFilePath & StrEncPwC & "&submit=submit"
    
    ' 폼 제출
    If IsPwC = True Then
        Call SubmitGoogleForm(FormURL1 & StrForm)
    Else
        Call SubmitGoogleForm(FormURL1 & StrForm)
        'Call SubmitGoogleForm(FormURL1 & StrEncForm)
    End If
    
    If IsPwC = False Then
        Call SubmitGoogleForm(FormURL2 & StrForm)
        'Call SubmitGoogleForm(FormURL3 & StrEncForm)
    End If

    If IsMandatory = True Then Exit Sub
    ' If UpdateTime = True Then Call AddOpenTime
End Sub

Function GetOpenTime() As Double
    On Error Resume Next
    Dim DetCel As Range, StrKey As String
    StrKey = GetMACAddress
    For Each DetCel In Sources.ListObjects("UserMaster").ListColumns(1).DataBodyRange
        If DetCel.value = StrKey Then GetOpenTime = DetCel.offset(0, 3).value: Exit Function
    Next DetCel
End Function

Sub AddOpenTime()
    On Error Resume Next
    Dim DetCel As Range, StrKey As String
    
    StrKey = GetMACAddress
    If IsPwC Then LastPwC.value = GetUserMSMail
    If Sources.ListObjects("UserMaster").DataBodyRange Is Nothing Then
        If Application.ScreenUpdating = False Then Exit Sub
        Sources.Cells(5, 3).value = Now
        Msg "등록되지 않은 사용자입니다." & vbLf & "사용자 등록 후 재실행해주세요.": Call EditUser: Exit Sub
    End If
    For Each DetCel In Sources.ListObjects("UserMaster").ListColumns(1).DataBodyRange
        If DetCel.value = StrKey Then DetCel.offset(0, 3).value = Now: GoTo GoExit
    Next DetCel
    Sources.Cells(5, 3).value = Now
    Msg "등록되지 않은 사용자입니다." & vbLf & "사용자 등록 후 재실행해주세요.": Call EditUser
GoExit:
    Set DetCel = Nothing
End Sub

Sub EditUser()
    On Error Resume Next
    Dim TblUser As ListObject, DetCel As Range
    IsBig3
    Set DetCel = Sources.ListObjects("UserMaster").ListColumns(1).DataBodyRange.Find(What:=GetMACAddress).offset(0, 1)
    If DetCel Is Nothing Then
        FrmUSER.Show
    Else
        SelectChoice DetCel.value & "님은 이미 등록된 사용자입니다." & vbLf & "기존 등록 정보를 수정하시겠습니까?"
        FrmUSER.Show
    End If
End Sub
Sub SelectChoice(ByVal StrMsg As String, Optional ByVal Opt As Long)
    Dim Slct As Long
    If Opt = 0 Then StrMsg = StrMsg & vbLf & vbLf & "작업은 취소할 수 없습니다."
    Slct = MsgBox(StrMsg, vbYesNo + vbQuestion, Title:="스마트리뷰어 " & StrSRVer)
    Select Case Slct
    Case vbNo: Call SpeedDown: End
    End Select
End Sub

Sub IsBig3()
    Dim StrMsg As String, StrMail As String
    StrMail = UCase(UserMail)
    If StrMail Like "*KPMG.COM" Then StrMsg = "삼정회계법인은": GoTo Big3
    If StrMail Like "*DELOITTE.COM" Then StrMsg = "안진회계법인은": GoTo Big3
    If StrMail Like "*EY.COM" Then StrMsg = "한영회계법인은": GoTo Big3
    Exit Sub
    
Big3:
    Call TimeToGoogle(True, False)
    StrMsg = StrMsg & vbLf & "Robotic Platform 서비스 대상 회사가 아닙니다."
    StrMsg = StrMsg & vbLf & "비권한자의 사용 로그를 집계합니다." & vbLf & "사용을 삼가주세요."
    Msg StrMsg: ThisWorkbook.Close False
    
End Sub

Function GetMACAddress() As String
    On Error Resume Next
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objItem In colItems
        GetMACAddress = objItem.MACAddress
        Exit Function
    Next
End Function

Function Encrypt_form(plaintext) As String
    Dim IV() As Byte
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set AES = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    AES.KeySize = 128
    AES.BlockSize = 128
    AES.Mode = 1 'CipherMode.CBC
    AES.padding = 2 'PaddingMode.PKCS7
    
    IV = GenerateIV(16) ' 16바이트 길이의 IV 생성
    AES.IV = IV
    
    aesKeyBytes = B64Decode("Y3VrZ2Jja2lyb2JvdGljIQ==")
    macKeyBytes = B64Decode("")
    Set aesEnc = AES.CreateEncryptor_2((aesKeyBytes), AES.IV)
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    macBytes = ComputeMAC(ConcatBytes(AES.IV, cipherBytes), macKeyBytes)

    ' IV와 암호문을 Base64로 인코딩
    encodedIV = B64Encode(IV)
    encodedCipher = B64Encode(cipherBytes)
    Encrypt_form = "('" & encodedIV & "', '" & encodedCipher & "')"

End Function

Public Function UserName() As String
    UserName = Application.UserName
End Function

Function ENCODEURL(varText As Variant, Optional blnEncode = True) As String
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        End With
    End If
    If blnEncode Then ENCODEURL = objHtmlfile.parentWindow.encode(varText)
    Set objHtmlfile = Nothing
End Function

Public Function UserMail() As String
    On Error Resume Next
    Select Case GetUserMSMail
    Case Is <> vbNullString: UserMail = GetUserMSMail
    Case vbNullString:
        'UserMail = GetUserOutlookMail
    End Select
    If UserMail = vbNullString Then UserMail = "N/A"
End Function

Public Function StrSRVer() As String
    StrSRVer = "PwC " & SRVer & "(" & RelVersion & ")" 'isinternal 삭제
    'Msg StrSRVer
    
End Function

Public Function RelDate() As Date
    RelDate = DateSerial(2024, 10, 16) '[update]
End Function
Public Function ExpDate() As Date
    ExpDate = DateSerial(2024, 12, 31) '현재 사용되지 않음.
End Function

Public Function CompanyName() As Range
    Set CompanyName = Sources.Cells(1, 1)
End Function

Public Function URLAbbName() As String
    URLAbbName = ENCODEURL(CompanyName.value)
End Function

Public Function URLFilePath() As String
    URLFilePath = ENCODEURL(FilePath1.value)
End Function

Public Function FilePath1() As Range
    Set FilePath1 = CompanyInfo.Cells(5, 1)
End Function

Public Function CompanyInfo() As Range
    Set CompanyInfo = Sources.Cells(1, 1).Resize(4, 5)
End Function

Public Function URLFilePathName() As String
    URLFilePathName = ENCODEURL(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
End Function
Sub SubmitGoogleForm(ByVal FormURL As String)
    On Error Resume Next
    Dim TicketInfo As MSXML2.ServerXMLHTTP60
    Dim headerName As String
    headerName = "Content-Type"
    Set TicketInfo = New ServerXMLHTTP60
    With TicketInfo
        .Open "POST", FormURL, False
        .setRequestHeader headerName, "application/x-www-form-urlencoded;charset=utf-8"
        .send
    End With
End Sub

Function GenerateIV(ByVal length As Long) As Byte()
    Dim hProv As LongPtr
    Dim IV() As Byte
    ReDim IV(length - 1)

    ' Crypto context 획득
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) Then
        ' IV 생성
        If CryptGenRandom(hProv, length, IV(0)) Then
            GenerateIV = IV
        End If
        CryptReleaseContext hProv, 0
    End If
End Function

Function B64Decode(b64Str)
    On Error Resume Next
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform")
    bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((bytes), 0, LenB(bytes))
End Function

Function ConcatBytes(a, b)
    On Error Resume Next
    Set mem = CreateObject("System.IO.MemoryStream")
    mem.SetLength (0)
    mem.Write (a), 0, LenB(a)
    mem.Write (b), 0, LenB(b)
    ConcatBytes = mem.ToArray()
End Function

Function ComputeMAC(msgBytes, keyBytes)
    On Error Resume Next
    Set Mac = CreateObject("System.Security.Cryptography.HMACSHA256")
    Mac.key = keyBytes
    ComputeMAC = Mac.ComputeHash_2((msgBytes))
End Function

Function B64Encode(bytes)
    On Error Resume Next
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform")
    BlockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(bytes) - 1 Step BlockSize
        length = min(BlockSize, LenB(bytes) - offset)
        b64Block = b64Enc.TransformFinalBlock((bytes), offset, length)
        result = result & utf8.GetString((b64Block))
    Next
    B64Encode = result
End Function

Function min(a, b)
    min = a
    If b < a Then min = b
End Function
Public Function ThisFont() As String
    ThisFont = ThisWorkbook.Styles("Normal").Font.Name
End Function

Public Function PwCOrg() As Long
    PwCOrg = RGB(208, 74, 2)
End Function
Public Function StrCopyRight() As String
    StrCopyRight = "ⓒ " & year(Now) & " Samil PwC. All rights reserved."
End Function
Sub FramePosition(ByVal Frame As Object)
    Dim sngLeft As Single, sngTop As Single
    With Frame
        .StartUpPosition = 0
        .Caption = "Smart_NPL " & StrSRVer
        Call ReturnPosition_CenterScreen(.Height, .Width, sngLeft, sngTop)
        .Left = sngLeft: .Top = sngTop
    End With
End Sub
Public Sub ReturnPosition_CenterScreen(ByVal sngHeight As Single, _
                                       ByVal sngWidth As Single, _
                                       ByRef sngLeft As Single, _
                                       ByRef sngTop As Single)
    Dim sngAppWidth As Single
    Dim sngAppHeight As Single
    Dim hWnd As Long
    Dim lreturn As Long
    Dim lpRect As udtRECT

    hWnd = Application.hWnd   'Used in Excel and Word
    lreturn = GetWindowRect(hWnd, lpRect)
    sngAppWidth = ConvertPixelsToPoints(lpRect.Right - lpRect.Left, "X")
    sngAppHeight = ConvertPixelsToPoints(lpRect.Bottom - lpRect.Top, "Y")
    sngLeft = ConvertPixelsToPoints(lpRect.Left, "X") + ((sngAppWidth - sngWidth) / 2)
    sngTop = ConvertPixelsToPoints(lpRect.Top, "Y") + ((sngAppHeight - sngHeight) / 2)
    
End Sub

Public Function ConvertPixelsToPoints(ByVal sngPixels As Single, _
                                      ByVal sXorY As String) As Single
    Dim hDC As Long
    hDC = GetDC(0)
    If sXorY = "X" Then
       ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 88))
    End If
    If sXorY = "Y" Then
       ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 90))
    End If
    Call ReleaseDC(0, hDC)
End Function

Sub Msg(ByVal StrMsg As String)
    MsgBox StrMsg, Title:="Smart_NPL " & StrSRVer
End Sub
