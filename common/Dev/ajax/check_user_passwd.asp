<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : check_user_passwd.asp / 회원비밀번호 체크
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Passwd : Passwd = FN_Req("Passwd","")
	Dim Rs
	Dim EncPasswd : EncPasswd = ""

	If Passwd <> "" Then
		EncPasswd = hex_sha1(Passwd)
	End If
			
	With objCmd
		.ActiveConnection = objDBCon
		.CommandType = 4
		.CommandText ="uspUser_Info"
		.Parameters.Refresh
		.Parameters("@FLAG").Value 	= "SIMP"
		.Parameters("@Seq").Value 	= Session("USeq")
			
		' ## 회원정보
		Set Rs = .Execute

		If Not Rs.Eof Then
			If EncPasswd = Rs("Passwd") Then
				Response.Write "true"
			Else
				Response.Write "false"
			End If
		Else
			response.Write "false"
		End If

		Rs.Close
		Set Rs = Nothing
	End With
		
%>
