<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 1. 권한 체크 (임시 세션 존재 여부 확인)
    If Session("TempASeq") = "" Or IsNull(Session("TempASeq")) Then
        Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='/admin/login.asp';</script>"
        Response.End
    End If

    Dim action, otpHelper, userSecret, qrUrl, otpCode, isVerify, checkRs, sqlCheck
    action = Request.Form("action")

    ' DB에 이미 Secret이 저장되어 있는지 재차 확인
    userSecret = ""
    Dim tables, tbl, qrySuccess
    tables = Array("tblAdmin", "Admin", "Member_Admin", "tbl_Admin")
    qrySuccess = False
    Dim currentTable
    currentTable = ""

    On Error Resume Next
    For Each tbl In tables
        If Not qrySuccess Then
            Err.Clear
            Set checkRs = objDbCon.Execute("SELECT AdmOtpSecret FROM " & tbl & " WHERE AdmSeq = " & Session("TempASeq"))
            If Err.Number = 0 Then
                If Not checkRs.Eof Then
                    userSecret = checkRs("AdmOtpSecret") & ""
                End If
                checkRs.Close
                Set checkRs = Nothing
                qrySuccess = True
                currentTable = tbl
            End If
        End If
    Next
    On Error GoTo 0

    ' 이미 등록된 사용자라면 검증 페이지로 리다이렉트
    If qrySuccess And userSecret <> "" Then
        Response.Redirect "/admin/otp_auth.asp"
        Response.End
    End If

    ' 2. 신규 Secret 생성 혹은 세션 보관
    If Session("NewOtpSecret") = "" Then
        On Error Resume Next
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifier")
        If Err.Number = 0 Then
            Session("NewOtpSecret") = otpHelper.GenerateSecret()
        Else
            Response.Write "<script>alert('OTP 컴포넌트 로드 실패. (DLL 등록 상태를 확인하십시오)\nError: " & Err.Description & "'); history.back();</script>"
            Response.End
        End If
        Set otpHelper = Nothing
        On Error GoTo 0
    End If

    userSecret = Session("NewOtpSecret")

    ' QR 코드 URL 구성 (Google Chart API 사용)
    Dim qrPayload, issuer, account
    issuer = "MPCJOB"
    account = Session("TempAID")
    qrPayload = "otpauth://totp/" & issuer & ":" & account & "?secret=" & userSecret & "&issuer=" & issuer
    
    ' URL Encoding 수행
    Dim encodedPayload
    encodedPayload = Server.URLEncode(qrPayload)
    qrUrl = "https://chart.googleapis.com/chart?chs=200x200&chld=M|0&cht=qr&chl=" & encodedPayload

    ' 3. 인증번호 검증 및 DB 등록
    If action = "register" Then
        otpCode = Trim(Request.Form("otp_code"))
        If Len(otpCode) <> 6 Or Not IsNumeric(otpCode) Then
            Response.Write "<script>alert('OTP 번호 6자리를 입력해주세요.'); history.back();</script>"
            Response.End
        End If

        isVerify = False
        On Error Resume Next
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifier")
        If Err.Number = 0 Then
            isVerify = otpHelper.VerifyOtp(userSecret, otpCode)
        End If
        Set otpHelper = Nothing
        On Error GoTo 0

        If isVerify Then
            ' DB에 Secret Key 영구 저장
            On Error Resume Next
            Dim updateSql, rowsAffected
            updateSql = "UPDATE " & currentTable & " SET AdmOtpSecret = '" & userSecret & "' WHERE AdmSeq = " & Session("TempASeq")
            objDbCon.Execute updateSql, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('DB 등록 중 오류가 발생했습니다. 관리자 컬럼 추가 여부를 확인하십시오.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            ' 로그인 최종 완료 처리
            Session("ASeq")   = Session("TempASeq")
            Session("AID")    = Session("TempAID")
            Session("AName")  = Session("TempAName")
            Session("ALevel") = Session("TempALevel")
            Session.Timeout = 180

            ' 임시 세션 제거
            Session("TempASeq")     = ""
            Session("TempAID")      = ""
            Session("TempAName")    = ""
            Session("TempALevel")   = ""
            Session("NewOtpSecret") = ""

            Response.Write "<script>alert('OTP 기기가 성공적으로 등록되었습니다.'); location.href='/admin/recruit/job_skin_list.asp';</script>"
            Response.End
        Else
            Response.Write "<script>alert('인증번호가 일치하지 않습니다. 다시 시도하십시오.'); history.back();</script>"
            Response.End
        End If
    End If
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>OTP 기기 등록 - HC 관리자</title>
    <link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
    <script type="text/javascript">
        $(function() {
            $("#otp_code").focus();
        });
    </script>
</head>
<body>
    <div id="wrap">
        <div class="login" style="height: 520px; width: 450px;">
            <div class="inner" style="width: 400px;">
                <h1><span style="font-size:24px; font-weight:bold; color:#333;">OTP 최초 기기 등록</span></h1>
                <p style="margin: 10px 0; color: #666; font-size:12px; line-height:1.6;">
                    1. 스마트폰에 <strong>Google Authenticator</strong> 앱을 설치합니다.<br/>
                    2. 앱에서 오른쪽 아래 (+) 버튼을 누른 후 <strong>QR 코드 스캔</strong>을 선택하여 아래 QR 코드를 스캔해주십시오.
                </p>
                <div style="text-align:center; margin: 15px 0;">
                    <img src="<%=qrUrl%>" alt="OTP QR Code" style="border:1px solid #ccc; padding:5px; background:#fff;" />
                    <div style="font-size:11px; color:#999; margin-top:5px;">수동 등록 코드: <strong style="color:#e056fd;"><%=userSecret%></strong></div>
                </div>
                <p style="margin: 10px 0; color: #666; font-size:12px; line-height:1.6;">
                    3. QR 등록 후 앱에 나타나는 6자리 인증번호를 아래에 입력하여 검증을 완료하십시오.
                </p>
                <div class="loginForm" style="margin-top:15px; text-align:center;">
                    <form id="frmOtp" name="frmOtp" action="otp_register.asp" method="post">
                        <input type="hidden" name="action" value="register" />
                        <label for="otp_code" style="width:80px; font-weight:bold; line-height:30px;">인증번호</label>
                        <input type="text" id="otp_code" name="otp_code" maxlength="6" style="width:120px; font-size:16px; text-align:center; letter-spacing:4px; height:24px;" autocomplete="off" />
                        <input type="submit" value="등록 및 완료" style="padding: 5px 15px; margin-left: 10px; background:#22a6b3; color:#fff; border:none; cursor:pointer;" />
                    </form>
                </div>
            </div>
            <p class="copy" style="width:100%;"><img src="/admin/images/copy_admin.jpg" alt="COPYRIGHT (C) 2009 MPC LTD., ALL RIGHTS RESERVED" /></p>
        </div>
    </div>
</body>
</html>
