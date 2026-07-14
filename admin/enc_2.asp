<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%

	Dim strFilter, Sql, RS
	Dim i
	i= 0


		'Sql = " SELECT UsrSeq, UsrID, Mobile FROM TBL_USER WHERE RegDate <= '2015-01-01' ORDER BY RegDate "
		'Sql = " SELECT UsrSeq, UsrID, Mobile FROM TBL_USER WHERE RegDate >= '2015-01-01' and RegDate <= '2016-01-01' ORDER BY RegDate "
		'Sql = " SELECT UsrSeq, UsrID, Mobile FROM TBL_USER WHERE RegDate >= '2016-01-01' and RegDate <= '2017-01-01' ORDER BY RegDate "
		'Sql = " SELECT UsrSeq, UsrID, Mobile FROM TBL_USER WHERE RegDate >= '2017-01-01' and RegDate <= '2018-01-01' ORDER BY RegDate "
		Sql = " SELECT UsrSeq, UsrID, Mobile FROM TBL_USER WHERE RegDate >= '2018-01-01' and RegDate <= '2019-01-01' ORDER BY RegDate "


		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 1 ' adCmdText
			.CommandText =Sql
			Set Rs = .Execute
			
		End With
		

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		If StrComp("OK", returnMsg)= 0 Then

			do while RS.eof = false    '레코드셋에 값이 있으면 계속 반복
				If Rs(2) <> "" Then 
					With objCmd
						.ActiveConnection = objDbCon
						'.CommandText = "Update TBL_USER Set UsrID = '"&oSeed.Encrypt(Rs(1))&"' Where UsrSeq ='"&Rs(0)&"'"
						Response.write ("Update TBL_USER Set Mobile= '"&oSeed.Encrypt(Rs(2))&"' Where UsrSeq ='"&Rs(0)&"'")& "<BR>"
						'.Execute
					End With
				End If 
				i = i + 1
			RS.moveNext
			Loop
			
			response.write "총 : "& i

			Rs.Close
			Set Rs = Nothing

		Set oSeed = Nothing
		End If 

%>