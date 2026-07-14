<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : memSecession_proc.asp / HC - 회원탈퇴 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim Result
	Dim UserID, UserName, Passwd
	Dim EncPasswd

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub

	Sub getParam()
	
		UserID 	= FN_Req("UserID","")
		UserName = FN_Req("UserName","")
		Passwd = FN_Req("Passwd","")
		
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
			.Parameters("@Flag").Value 		= "OUT"
			.Parameters("@Seq").Value 		= Session("USeq")
			.Parameters("@Mobile").Value 	= oSeed.Encrypt(UserID)
			.Parameters("@UsrName").Value 	= UserName
			.Parameters("@Passwd").Value 	= EncPasswd

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = 99
			End If
			
		End With
		
		Set Rs = Nothing

		Set oSeed = Nothing
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
	
		If Result = 0 Then
			strMsg = "정상적으로 회원탈퇴 처리되었습니다."
			strScript = "window.location.href=""http://www.mpcjob.co.kr/membership/logout.asp"";"
		ElseIf Result = 1 Then
			strMsg = "입력해주신 정보가 회원정보와 일치하지 않습니다."
			strScript = "history.back();"
		Else 
			strMsg = "탈퇴처리 도중 오류가 발생하였습니다."
			strScript = "history.back();"
		End If
		
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>회원 정보 처리 - HC 관리자</title>
			<script type="text/javascript">
			<!--
				alert("<%=strMsg%>");
				<%=strScript%>;
			//-->
			</script>
		</head>
		<body>
		</body>
	</html>
	<%
	End Sub
%>
