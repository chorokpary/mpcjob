<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Function/FnAuditLog.asp" -->
<!-- #include virtual="/admin/siteconf/iplimit/ip_check.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : login_proc_otp.asp / MPCJOB - 관리자 : 로그인 처리 (OTP 테스트용)
' 
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>로그인 - HC 관리자 (OTP 테스트)</title>
<% 
	Dim strID, strPWD, Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
	End Sub
	
	Sub getParam()
	
		strID 		= Trim(FN_Req("txtID",""))
		strPWD 		= Trim(FN_Req("txtPWD",""))
		
		If strID = "" Or strPWD = "" Then 		Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","BACK")

		strPWD 		= hex_sha1(strPWD)
		
	End Sub
	
	Sub setProc()
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspAdm_Auth"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "LOGIN"
			
			.Parameters("@AdminID").Value 		= strID
			.Parameters("@AdminPwd").Value 		= strPWD
			.Parameters("@UserIP").Value 		= Request.ServerVariables("REMOTE_ADDR")

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = "ERROR"
			End If
			
			If Result = "SUCC" Then
				Set Rs = Rs.NextRecordSet
				
				' 임시 세션 발급 (로그인 1단계 성공 상태)
				Session("TempASeq")   = Rs("AdmSeq")
				Session("TempAID")    = Rs("AdmID")
				Session("TempAName")  = Rs("AdmName")
				Session("TempALevel") = Rs("AdmLevel")

				Session.Timeout = 30 ' OTP 입력 제한 시간 (30분)

				' DB에서 사용자의 OTP Secret 존재 여부 체크 (후보 테이블 순차 조회)
				Dim dbSecretVal, checkRs, sqlCheck
				dbSecretVal = ""
				
				Dim tables, tbl, qrySuccess
				tables = Array("tblAdmin", "Admin", "Member_Admin", "tbl_Admin")
				qrySuccess = False

				On Error Resume Next
				For Each tbl In tables
					If Not qrySuccess Then
						Err.Clear
						Set checkRs = objDbCon.Execute("SELECT AdmOtpSecret FROM " & tbl & " WHERE AdmSeq = " & Rs("AdmSeq"))
						If Err.Number = 0 Then
							If Not checkRs.Eof Then
								dbSecretVal = checkRs("AdmOtpSecret") & ""
							End If
							checkRs.Close
							Set checkRs = Nothing
							qrySuccess = True
						End If
					End If
				Next
				On Error GoTo 0
				
				If qrySuccess And dbSecretVal <> "" Then
					' [감사 로그] 1단계 인증 성공 기록 (OTP 대기)
					Call AddAuditLog("LOGIN_OTP_WAIT", "000000", Rs("AdmSeq"), "관리자 1차 로그인 성공 - OTP 인증 대기 (ID: " & strID & ", 이름: " & Rs("AdmName") & ")")
					
					' OTP가 이미 등록되어 있는 경우 -> OTP 인증 페이지로 리다이렉트
					%><meta http-equiv="refresh" content="0; url=/admin/otp_auth.asp"><%
				Else
					' [감사 로그] 1단계 인증 성공 기록 (OTP 미등록)
					Call AddAuditLog("LOGIN_OTP_WAIT", "000000", Rs("AdmSeq"), "관리자 1차 로그인 성공 - OTP 신규 등록 대기 (ID: " & strID & ", 이름: " & Rs("AdmName") & ")")
					
					' OTP 등록이 필요한 경우 -> OTP 최초 기기 등록 페이지로 리다이렉트
					%><meta http-equiv="refresh" content="0; url=/admin/otp_register.asp"><%
				End If
				Response.End
			ElseIf Result = "LOCK" Then
				' [감사 로그] 로그인 잠금 기록
				Call AddAuditLog("LOGIN_FAIL", "000000", "", "로그인 실패 - 계정 잠금상태 (ID: " & strID & ")")
				Call SB_ReturnErr("로그인 실패 횟수 초과 되었습니다. 담당자에게 연락 바랍니다.","BACK")
			ElseIf Result = "FAIL" Then
				' [감사 로그] 로그인 실패 기록
				Call AddAuditLog("LOGIN_FAIL", "000000", "", "로그인 실패 - 접속정보 불일치 (ID: " & strID & ")")
				Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","BACK")
			Else
				' [감사 로그] 시스템 오류 기록
				Call AddAuditLog("LOGIN_FAIL", "000000", "", "로그인 실패 - 내부 시스템 오류 (ID: " & strID & ")")
				Call SB_ReturnErr("처리도중 오류가 발생하였습니다. 다시 시도해 주시기 바랍니다.","BACK")
			End If
		End With
		
	End Sub
%>
</head>
<body>
</body>
</html>
