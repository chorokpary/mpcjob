<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : login_proc.asp / MPCJOB - 관리자 : 로그인 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>로그인 - HC 관리자</title>
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
				
				Session("ASeq") = Rs("AdmSeq")
				Session("AID")	= Rs("AdmID")
				Session("AName") = Rs("AdmName")
				Session("ALevel") = Rs("AdmLevel")

				Session.Timeout = 180
				
				%><meta http-equiv="refresh" content="0; url=http://www.mpcjob.co.kr/admin/recruit/job_skin_list.asp"><%
				Response.End
			ElseIf Result = "LOCK" Then
				Call SB_ReturnErr("로그인 실패 횟수 초과 되었습니다. 담당자에게 연락 바랍니다.","BACK")
			ElseIf Result = "FAIL" Then
				Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","BACK")
			Else
				Call SB_ReturnErr("처리도중 오류가 발생하였습니다. 다시 시도해 주시기 바랍니다.","BACK")
			End If
		End With
		
	End Sub
%>
</head>
<body>
</body>
</html>