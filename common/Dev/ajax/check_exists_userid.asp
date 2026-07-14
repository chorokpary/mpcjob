<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : check_exists_userid.asp / 회원아이디(=모바일번호) 중복 체크
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-12 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Mobile : Mobile = FN_Req("Mobile","")
	Dim Rs

	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 
	
	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)


	With objCmd
		.ActiveConnection = objDBCon
		.CommandType = 4
		.CommandText ="uspUser_Info"
		.Parameters.Refresh
		.Parameters("@FLAG").Value 		= "EXIST"
		.Parameters("@UsrID").Value 	= oSeed.Encrypt(Mobile)
			
		' ## 회원정보
		Set Rs = .Execute

		Set oSeed = Nothing

		If Not Rs.Eof Then
			Response.Write "false"
		Else
			response.Write "true"
		End If

		Rs.Close
		Set Rs = Nothing
	End With
		
%>
