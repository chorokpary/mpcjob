<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : find_proc.asp / MPCJOB - 비밀번호 조회 처리 (Ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-25 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim strID, strName
	Dim UsrSeq, Passwd, Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
	End Sub
	
	Sub getParam()
	
		strID 		= FN_Req("txtID","")
		strName 	= FN_Req("txtName","")
		
	End Sub
	
	Sub setProc()
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "FIND"
			
			.Parameters("@UsrID").Value 		= strID
			.Parameters("@UsrName").Value 		= strName

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = "SUCC"
				UsrSeq = Rs("UsrSeq")
			Else
				Result = "FAIL"
				UsrSeq = 0
			End If
			
		End With
		
		Set Rs = Nothing

		If Result = "SUCC" Then
			' 아이디가 일치할 경우 새 비밀번호를 업데이트 하고 SMS 발송
			Dim encPwd
			
			Passwd = Fn_EncryptData("NEWPWD")
			Passwd = LEFT(Passwd,6)
			Passwd = UCase(Passwd)
				
			encPwd = hex_sha1(Passwd)

			' 새비번 UPDATE
			With objCmd
				.ActiveConnection = objDbCon
				.CommandType = 4
				.CommandText ="uspUser_Proc"
				.Parameters.Refresh
				.Parameters("@Flag").Value 		= "PWD_C"
				
				.Parameters("@Seq").Value 		= UsrSeq
				.Parameters("@Passwd").Value 	= encPwd
	
				.Execute
				
				Set Rs = .Execute
				
				If Not Rs.Eof Then
					If Rs("Result") = "0" Then
						'SMS발송 (Passwd)
						Call SB_SendSMS(strID, 0, "S","[비밀번호찾기] " & strID & " 회원", "#MPCJOB# 임시 비밀번호를 전송합니다. [" & Passwd & "]",0,"","MEM",0,"")
					Else
						Result = "FAIL"
					End If
				End If
				
			End With

		End If
		
		Set Rs = Nothing
		
		Response.Write Result
		Response.End
		
	End Sub
%>
