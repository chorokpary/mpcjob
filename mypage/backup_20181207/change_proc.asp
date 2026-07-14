<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : change_proc.asp / HC - 핸드폰번호 OR 비밀번호 변경 처리 (Ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim Flag, Result
	Dim Mobile, ReMobile
	Dim Passwd, RePasswd, EncPasswd

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub

	Sub getParam()
	
		Flag 	= FN_Req("Flag","")
		
		Mobile 	= FN_Req("New_Mobile_1","") & "-" & FN_Req("New_Mobile_2","") & "-" & FN_Req("New_Mobile_3","")
		ReMobile = FN_Req("Re_Mobile_1","") & "-" & FN_Req("Re_Mobile_2","") & "-" & FN_Req("Re_Mobile_3","")
		Passwd = FN_Req("New_Passwd","")
		
		If Passwd <> "" Then
			EncPasswd = hex_sha1(Passwd)
		End If
		
	End Sub
	
	Sub setProc()

		set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters("@Seq").Value 		= Session("USeq")
			.Parameters("@Passwd").Value 	= EncPasswd
			.Parameters("@Mobile").Value 	= oSeed.Encrypt(Mobile)

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = 99
			End If
			
		End With
		
		Set Rs = Nothing

		If Result = "0" Then
			Session("UType") = 1
		End If
		
		Set oSeed = Nothing

	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			If Flag = "MOBILE_C" Then
				Session("UID")	= Mobile
			End If
			Response.Write "SUCC"
		ElseIf Result = 1 Then
			Response.Write "EXISTS"
		Else 
			Response.Write "FAIL"
		End If
		Response.End

	End Sub
%>
