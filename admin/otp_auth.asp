<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 1. 권한 체크 (임시 세션 존재 여부 확인)
    If Session("TempASeq") = "" Or IsNull(Session("TempASeq")) Then
        Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='/admin/login.asp';</script>"
        Response.End
    End If

    Dim action, otpCode, isVerify, userSecret, checkRs, sqlCheck
    action = Request.Form("action")

    ' 2. DB에서 사용자의 OTP Secret 조회 (후보 테이블 순차 조회)
    userSecret = ""
    Dim tables, tbl, qrySuccess
    tables = Array("tblAdmin", "Admin", "Member_Admin", "tbl_Admin")
    qrySuccess = False

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
            End If
        End If
    Next
    On Error GoTo 0

    ' 만약 Secret Key가 없다면 등록 페이지로 강제 이동
    If qrySuccess And userSecret = "" Then
        Response.Redirect "/admin/otp_register.asp"
        Response.End
    End If

    ' 3. OTP 검증 처리
    If action = "verify" Then
        otpCode = Trim(Request.Form("otp_code"))
        If Len(otpCode) <> 6 Or Not IsNumeric(otpCode) Then
            Response.Write "<script>alert('OTP 번호 6자리를 입력해주세요.'); history.back();</script>"
            Response.End
        End If

        isVerify = False
        On Error Resume Next
        Dim otpHelper
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifier")
        If Err.Number = 0 Then
            isVerify = otpHelper.VerifyOtp(userSecret, otpCode)
        Else
            ' COM DLL이 등록되지 않았거나 오류 시 안내
            Response.Write "<script>alert('OTP 컴포넌트 오류가 발생했습니다. (DLL 등록 상태 확인 필요)\nError: " & Err.Description & "'); history.back();</script>"
            Response.End
        End If
        Set otpHelper = Nothing
        On Error GoTo 0

        If isVerify Then
            ' 로그인 최종 완료 처리
            Session("ASeq")   = Session("TempASeq")
            Session("AID")    = Session("TempAID")
            Session("AName")  = Session("TempAName")
            Session("ALevel") = Session("TempALevel")
            Session.Timeout = 180

            ' 임시 세션 제거
            Session("TempASeq")   = ""
            Session("TempAID")    = ""
            Session("TempAName")  = ""
            Session("TempALevel") = ""

            Response.Write "<script>location.href='/admin/recruit/job_skin_list.asp';</script>"
            Response.End
        Else
            Response.Write "<script>alert('OTP 인증번호가 일치하지 않습니다.'); history.back();</script>"
            Response.End
        End If
    End If
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>OTP 인증 - HC 관리자</title>
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
        <div class="login" style="height: 380px;">
            <div class="inner">
                <h1><span style="font-size:24px; font-weight:bold; color:#333;">OTP 추가 인증</span></h1>
                <p style="margin: 15px 0; color: #666; font-size:12px; line-height:1.6;">
                    스마트폰의 Google Authenticator 또는 Microsoft Authenticator 앱에 표시된 6자리 인증 번호를 입력해주십시오.
                </p>
                <div class="loginForm" style="margin-top:20px;">
                    <form id="frmOtp" name="frmOtp" action="otp_auth.asp" method="post">
                        <input type="hidden" name="action" value="verify" />
                        <label for="otp_code" style="width:80px; font-weight:bold; line-height:30px;">인증번호</label>
                        <input type="text" id="otp_code" name="otp_code" maxlength="6" style="width:120px; font-size:16px; text-align:center; letter-spacing:4px; height:24px;" autocomplete="off" />
                        <input type="submit" value="인증" style="padding: 5px 15px; margin-left: 10px; background:#47596b; color:#fff; border:none; cursor:pointer;" />
                    </form>
                </div>
            </div>
            <p class="copy"><img src="/admin/images/copy_admin.jpg" alt="COPYRIGHT (C) 2009 MPC LTD., ALL RIGHTS RESERVED" /></p>
        </div>
    </div>
</body>
</html>
