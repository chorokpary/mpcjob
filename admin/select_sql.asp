<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%

	Dim strFilter, Sql, RS
	Dim i
	i= 0

		Sql = " SELECT " & _
			  "     m.UsrSeq, " & _
			  "     (SELECT usrName FROM tbl_user WHERE usrseq = m.UsrSeq) AS name, " & _
			  "     (SELECT UsrID FROM tbl_user WHERE usrseq = m.UsrSeq) AS userID, " & _
			  "     m.Memo, " & _
			  "     m.RegDate " & _
			  " FROM TBL_USER_MEMO AS m " & _
			  " WHERE m.RegDate >= '2023-01-01' " & _
			  " ORDER BY m.RegDate DESC "


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
				If Rs(1) <> "" Then 
					With objCmd
						.ActiveConnection = objDbCon
						Response.write ( Rs(1) &"|"& oSeed.Decrypt(Rs(2)) &"|"& Rs(3) &"|"& Rs(4) )& "<BR>"
					End With
				End If 
				i = i + 1
			RS.moveNext
			Loop
			
			Rs.Close
			Set Rs = Nothing

		Set oSeed = Nothing
		End If 

%>