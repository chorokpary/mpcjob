<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 1. 권한 체크 (임시 세션 존재 여부 확인)
    If Session("TempASeq") = "" Or IsNull(Session("TempASeq")) Then
        Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='/admin/login_otp.asp';</script>"
        Response.End
    End If

    Dim action, otpCode, isVerify, userSecret, checkRs, sqlCheck
    action = Request.Form("action")

    ' 2. DB에서 사용자의 OTP Secret 조회 및 정규식 필터링 (눈에 안 보이는 공백/가비지 방지)
    userSecret = ""
    Dim currentTable, qrySuccess
    currentTable = "TBL_ADM"
    qrySuccess = False

    On Error Resume Next
    Set checkRs = objDbCon.Execute("SELECT AdmOtpSecret FROM " & currentTable & " WHERE AdmSeq = " & Session("TempASeq"))
    If Err.Number = 0 Then
        If Not checkRs.Eof Then
            Dim rawSecret, regEx
            rawSecret = checkRs("AdmOtpSecret") & ""
            
            ' Base32에 유효한 알파벳(A-Z) 및 숫자(2-7)만 남기고 모든 공백, 개행, 특수문자 제거
            Set regEx = New RegExp
            regEx.Pattern = "[^a-zA-Z2-7]"
            regEx.IgnoreCase = True
            regEx.Global = True
            userSecret = regEx.Replace(rawSecret, "")
        End If
        checkRs.Close
        Set checkRs = Nothing
        qrySuccess = True
    End If
    On Error GoTo 0

    ' 3. OTP 재등록(초기화) 처리
    If action = "reset" Then
        On Error Resume Next
        objDbCon.Execute("UPDATE " & currentTable & " SET AdmOtpSecret = NULL WHERE AdmSeq = " & Session("TempASeq"))
        On Error GoTo 0
        Session("NewOtpSecret") = ""
        Response.Redirect "/admin/otp_register.asp"
        Response.End
    End If

    ' 만약 Secret Key가 없다면 등록 페이지로 강제 이동
    If qrySuccess And userSecret = "" Then
        Response.Redirect "/admin/otp_register.asp"
        Response.End
    End If

    ' 4. OTP 검증 처리
    If action = "verify" Then
        otpCode = Trim(Request.Form("otp_code"))
        If Len(otpCode) <> 6 Or Not IsNumeric(otpCode) Then
            Response.Write "<script>alert('OTP 번호 6자리를 입력해주세요.'); history.back();</script>"
            Response.End
        End If

        isVerify = False
        On Error Resume Next
        Dim otpHelper
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifierV2")
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
    <style type="text/css">
        body {
            margin: 0;
            padding: 0;
            background: #f1f3f6;
            font-family: 'Malgun Gothic', 'Dotum', sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        .login-card {
            background: #ffffff;
            width: 400px;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.05);
            border: 1px solid #eef2f5;
            box-sizing: border-box;
        }
        .login-header {
            text-align: center;
            margin-bottom: 25px;
        }
        .login-header h1 {
            font-size: 24px;
            color: #2c3e50;
            margin: 0 0 10px 0;
            font-weight: 700;
            letter-spacing: -0.5px;
        }
        .login-header p {
            font-size: 13px;
            color: #00529c;
            margin: 0;
            font-weight: bold;
        }
        .instruction-box {
            background: #f8f9fa;
            border-left: 4px solid #00529c;
            padding: 12px 15px;
            font-size: 12px;
            color: #4f5f6f;
            line-height: 1.6;
            border-radius: 4px;
            margin-bottom: 25px;
        }
        .form-group {
            margin-bottom: 25px;
            text-align: center;
        }
        .form-group label {
            display: block;
            font-size: 14px;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .form-control {
            width: 150px;
            height: 42px;
            padding: 5px;
            border: 1px solid #dcdde1;
            border-radius: 6px;
            box-sizing: border-box;
            font-size: 20px;
            font-weight: bold;
            color: #00529c;
            text-align: center;
            letter-spacing: 6px;
            transition: all 0.3s ease;
        }
        .form-control:focus {
            border-color: #00529c;
            outline: none;
            box-shadow: 0 0 0 3px rgba(0, 82, 156, 0.15);
        }
        .submit-btn {
            width: 100%;
            height: 45px;
            background: #00529c;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-size: 15px;
            font-weight: bold;
            cursor: pointer;
            transition: background 0.3s ease, transform 0.1s ease;
            margin-top: 10px;
        }
        .submit-btn:hover {
            background: #003d75;
        }
        .submit-btn:active {
            transform: scale(0.98);
        }
        .reset-btn {
            width: 100%;
            height: 40px;
            background: #ffffff;
            color: #7f8c8d;
            border: 1px solid #dcdde1;
            border-radius: 6px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-top: 10px;
        }
        .reset-btn:hover {
            background: #fff5f5;
            color: #e74c3c;
            border-color: #e74c3c;
        }
        .reset-btn:active {
            transform: scale(0.98);
        }
        .login-copy {
            text-align: center;
            margin-top: 30px;
            font-size: 11px;
            color: #bdc3c7;
        }
    </style>
    <script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
    <script type="text/javascript">
        $(function() {
            $("#otp_code").focus();
        });
        function fnResetOtp() {
            if (confirm("OTP 앱을 삭제했거나 기기를 변경하셨습니까?\n\n기존 OTP 등록 정보를 초기화하고 새로 QR 코드를 스캔하여 재등록합니다.\n진행하시겠습니까?")) {
                $("#hdnAction").val("reset");
                $("#frmOtp").submit();
            }
        }
    </script>
</head>
<body>
    <div class="login-card">
        <div class="login-header">
            <h1>HC 관리자 시스템</h1>
            <p>OTP 추가 인증</p>
        </div>
        
        <div class="instruction-box">
            스마트폰의 <strong>Google OTP</strong> 또는 <strong>Microsoft Authenticator</strong> 앱에 표시된 6자리 인증 번호를 입력해주십시오.
        </div>

        <form id="frmOtp" name="frmOtp" action="otp_auth.asp" method="post">
            <input type="hidden" id="hdnAction" name="action" value="verify" />
            <div class="form-group">
                <label for="otp_code">인증번호 입력</label>
                <input type="text" id="otp_code" name="otp_code" class="form-control" maxlength="6" autocomplete="off" />
            </div>
            <button type="submit" class="submit-btn">인증</button>
            <button type="button" class="reset-btn" onclick="fnResetOtp();">OTP 앱 삭제 / 기기 변경시 재등록</button>
        </form>
        
        <div class="login-copy">
            COPYRIGHT © MPC Plus INC. ALL RIGHTS RESERVED.
        </div>
    </div>
</body>
</html>
