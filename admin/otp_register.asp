<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 1. 권한 체크 (임시 세션 존재 여부 확인)
    If Session("TempASeq") = "" Or IsNull(Session("TempASeq")) Then
        Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='/admin/login_otp.asp';</script>"
        Response.End
    End If

    Dim action, otpHelper, userSecret, qrUrl, otpCode, isVerify, checkRs, sqlCheck
    action = Request.Form("action")

    ' DB에 이미 Secret이 저장되어 있는지 재차 확인
    userSecret = ""
    Dim currentTable, qrySuccess
    currentTable = "TBL_ADM"
    qrySuccess = False

    On Error Resume Next
    Set checkRs = objDbCon.Execute("SELECT AdmOtpSecret FROM " & currentTable & " WHERE AdmSeq = " & Session("TempASeq"))
    If Err.Number = 0 Then
        If Not checkRs.Eof Then
            userSecret = checkRs("AdmOtpSecret") & ""
        End If
        checkRs.Close
        Set checkRs = Nothing
        qrySuccess = True
    End If
    On Error GoTo 0

    ' 이미 등록된 사용자라면 검증 페이지로 리다이렉트
    If qrySuccess And userSecret <> "" Then
        Response.Redirect "/admin/otp_auth.asp"
        Response.End
    End If

    ' 2. 신규 Secret 생성 혹은 세션 보관
    If Session("NewOtpSecret") = "" Then
        On Error Resume Next
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifierV2")
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
        Dim errNum, errDesc
        Set otpHelper = Server.CreateObject("Mpcjob.Otp.OtpVerifierV2")
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum = 0 Then
            isVerify = otpHelper.VerifyOtp(userSecret, otpCode)
            Set otpHelper = Nothing
        Else
            Response.Write "<script>alert('V2 개체 생성 실패! COM 등록을 다시 확인해주세요.\nError: " & errNum & " - " & errDesc & "'); history.back();</script>"
            Response.End
        End If

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
            width: 450px;
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
            margin-bottom: 20px;
        }
        .qr-area {
            text-align: center;
            margin: 20px 0;
            padding: 15px;
            background: #fcfcfc;
            border: 1px dashed #dcdde1;
            border-radius: 8px;
        }
        .qr-wrapper {
            display: inline-block;
            padding: 8px;
            background: #ffffff;
            border: 1px solid #e1e8ed;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }
        .secret-code {
            font-size: 12px;
            color: #7f8c8d;
            margin-top: 10px;
        }
        .secret-code strong {
            color: #e74c3c;
            font-size: 13px;
            font-family: Consolas, Monaco, monospace;
        }
        .form-group {
            margin-bottom: 20px;
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
    </script>
</head>
<body>
    <div class="login-card">
        <div class="login-header">
            <h1>HC 관리자 시스템</h1>
            <p>OTP 최초 기기 등록</p>
        </div>
        
        <div class="instruction-box">
            1. 스마트폰에 <strong>Google OTP</strong> 또는 <strong>Microsoft Authenticator</strong> 앱을 설치합니다.<br/>
            2. 앱에서 추가(+) 버튼을 누른 후 <strong>QR 코드 스캔</strong>을 선택하여 아래 QR 코드를 스캔해주십시오.
        </div>

        <div class="qr-area">
            <div class="qr-wrapper">
                <img src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=<%=Server.URLEncode(qrPayload)%>&t=<%=Server.URLEncode(Session("TempASeq") & "_" & Timer())%>" alt="OTP QR Code" style="width: 150px; height: 150px; display: block;" />
            </div>
            <div class="secret-code">
                수동 등록 코드: <strong><%=userSecret%></strong>
            </div>
        </div>

        <div class="instruction-box" style="margin-bottom: 15px; border-left-color: #27ae60;">
            3. 앱에 표시된 6자리 인증번호를 아래에 입력하여 등록을 완료해주십시오.
        </div>

        <form id="frmOtp" name="frmOtp" action="otp_register.asp" method="post">
            <input type="hidden" name="action" value="register" />
            <div class="form-group">
                <label for="otp_code">인증번호 입력</label>
                <input type="text" id="otp_code" name="otp_code" class="form-control" maxlength="6" autocomplete="off" />
            </div>
            <button type="submit" class="submit-btn">등록 및 완료</button>
        </form>
        
        <div class="login-copy">
            COPYRIGHT © MPC Plus INC. ALL RIGHTS RESERVED.
        </div>
    </div>
</body>
</html>
