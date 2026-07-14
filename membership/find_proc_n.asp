<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : find_proc.asp / HC - 비밀번호 조회 처리 (Ajax)
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

	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
	End Sub
	
	Sub getParam()
	
		strID 		= FN_Req("txtID","")
		strName 	= FN_Req("txtName","")

		If Session("ASeq") = "" Then 
			Session("ASeq") = 0
		End If 
		
	End Sub


	Function random_tmp_str_1()
		  Dim str, strlen, r, i, ds, serialCode '사용되는 변수를 선언
		  str = "!@#$%^&*()<>" '랜덤으로 사용될 문자 또는 숫자
		  strlen = 2 '랜덤으로 출력될 값의 자릿수 ex)해당 구문에서 10자리의 랜덤 값 출력
		   Randomize '랜덤 초기화
		   For i = 1 To strlen '위에 선언된 strlen만큼 랜덤 코드 생성
			r = Int((14 - 1 + 1) * Rnd + 1)  ' 36은 str의 문자갯수
			serialCode = serialCode + Mid(str,r,1)
		   Next
		   random_tmp_str_1 = serialCode
	End Function

	Function random_tmp_str_2()
		 Dim str, strlen, r, i, ds, serialCode '사용되는 변수를 선언
		  str = "123456789" '랜덤으로 사용될 문자 또는 숫자
		  strlen = 2 
		   Randomize '랜덤 초기화
		   For i = 1 To strlen 
			r = Int((9 - 1 + 1) * Rnd + 1)  
			serialCode = serialCode + Mid(str,r,1)
		   Next
		   random_tmp_str_2 = serialCode
	End Function


	Sub setProc()
		Dim Rs

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "FIND"
			
			.Parameters("@UsrID").Value 		= oSeed.Encrypt(strID)
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

		Set oSeed = Nothing

		If Result = "SUCC" Then
			' 아이디가 일치할 경우 새 비밀번호를 업데이트 하고 SMS 발송
			Dim encPwd
			
			Passwd = Fn_EncryptData("NEWPWD")
			Passwd = LEFT(Passwd,4)
			Passwd = UCase(Passwd)
			Passwd = Left(Passwd,2) & random_tmp_str_1() & random_tmp_str_2() & Right(Passwd,2)
				
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
						Call SB_Send_Emfo(Replace(strID,"-",""), 0, "SMS","[비밀번호찾기] " & strID & " 회원", "#MPCJOB# 임시 비밀번호를 전송합니다. [" & Passwd & "]", "N", "" ,"" ,"", "MEM", 0, "N", "", "")
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
