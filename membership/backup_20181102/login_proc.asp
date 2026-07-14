<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : login_proc.asp / HC - 로그인 처리 (Ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim strID, strPWD, chkSaveID, Result
	Dim NextUrl, PageFlag

	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
	End Sub
	
	Sub getParam()
	
		strID 		= Trim(FN_Req("txtID",""))
		strPWD 		= Trim(FN_Req("txtPWD",""))

		If strID = "" Or strPWD = "" Then 		Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","/membership/login.asp")

		strPWD 		= hex_sha1(strPWD)
		chkSaveID 	= FN_Req("chkSaveID","N")
		PageFlag 	= FN_Req("PageFlag","FRONT")
		NextUrl 	= FN_Req("NextUrl","/")
		
	End Sub
	
	Sub setProc()

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Auth"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "LOGIN"
			
			.Parameters("@UserID").Value 		= oSeed.Encrypt(strID)
			.Parameters("@UserPwd").Value 		= strPWD
			.Parameters("@UserIP").Value 		= Request.ServerVariables("REMOTE_ADDR")

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = "ERROR"
			End If

			
			
			If Result = "SUCC" Then
				Set Rs = Rs.NextRecordSet

				Session("USeq") = Rs("UsrSeq")
				Session("UID")	= oSeed.Decrypt(Rs("Mobile"))
				Session("UName") = Rs("UsrName")
				Session("UType") = Rs("UsrType")
				
				' 아이디 저장 확인
				If chkSaveID = "Y" Then
					Response.Cookies("UserID") = strID
					Response.Cookies("UserID").expires = Date + 100
				Else
					Response.Cookies("UserID") = ""
					Response.Cookies("UserID").expires = Date - 1
				End If
				
				If PageFlag = "FRONT" Then
					response.write Result
				ElseIf PageFlag = "ADM" Then
					%><meta http-equiv="refresh" content="0; url=/"><%
				End If
				Response.End
			ElseIf Result = "LOCK" Then
				'Call SB_ReturnErr("로그인 실패 횟수 초과 되었습니다. 담당자에게 연락 바랍니다.","SELF")

				response.write Result
				Response.End

			ElseIf Result = "FAIL" Then
			'	Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","SELF")

				If PageFlag = "FRONT" Then
					response.write Result
					Response.End
				ElseIf PageFlag = "ADM" Then
					Call SB_ReturnErr("접속 정보가 올바르지 않습니다.","/membership/login.asp")
				End If
			Else
				Call SB_ReturnErr("처리도중 오류가 발생하였습니다. 다시 시도해 주시기 바랍니다.","SELF")
			End If


			Set oSeed = Nothing

		End With
		
	End Sub
%>
