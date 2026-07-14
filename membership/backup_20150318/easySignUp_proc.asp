<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : easySignUp_proc.asp / 간편가입 처리 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4m / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, PageFlag, Seq
	Dim UsrName, Birthday, Gender, Mobile, HopePart, AddrSido, AddrGugun, Addr, EzCareer, RecrSeq
	Dim Passwd
	
	Dim Result
	
	Call Main()
	
	Sub Main()
	
		Flag 	= FN_Req("Flag","")
		PageFlag = FN_Req("PageFlag","")
		UsrName 	= FN_Req("UsrName","")
		Birthday 	= FN_Req("Birthday_Y","") & "-" & FN_Req("Birthday_M","") & "-" & FN_Req("Birthday_D","")
		Gender 	= FN_Req("Gender","")
		Mobile 	= FN_Req("Mobile_1","") & "-" & FN_Req("Mobile_2","") & "-" & FN_Req("Mobile_3","")
		HopePart 	= FN_Req("HopePart","")
		AddrSido 	= FN_Req("AddrSido","")
		AddrGugun = FN_Req("AddrGugun","")
		Passwd 	= FN_Req("Passwd","")
		EzCareer = FN_Req("EzCareer","")
		RecrSeq 	= FN_Req("RecrSeq","0")
		
		If Passwd <> "" Then		Passwd = hex_sha1(Passwd)

		Addr = AddrSido & " " & AddrGugun
		
		If Not IsNumeric(RecrSeq) Then RecrSeq = 0
		
		Call setProc()
		Call setResult()
	End Sub

	Sub setProc()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters("@Seq").Value 			= 0
			
			.Parameters("@Passwd").Value 		= Passwd
			.Parameters("@UsrName").Value 		= UsrName
			.Parameters("@Birthday").Value 		= Birthday
			.Parameters("@Gender").Value 		= Gender
			.Parameters("@Mobile").Value 		= Mobile
			.Parameters("@HopePart").Value 		= HopePart
			.Parameters("@AddrSido").Value 		= AddrSido
			.Parameters("@AddrGugun").Value 	= AddrGugun
			.Parameters("@EzCareer").Value 		= EzCareer

			.Parameters("@RecrSeq").Value 		= RecrSeq
			
			Set Rs = .Execute

			'Result = .Parameters("@Result")
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = 99
			End If
					
			If Flag = "ADD" And Result = 0 Then
			
				Set Rs = Rs.NextRecordSet()
				
				If Not Rs.Eof Then
					Session("USeq") = Rs("UsrSeq")
					Session("UID")	= Rs("Mobile")
					Session("UName") = Rs("UsrName")
					Session("UType") = Rs("UsrType")
					
					Seq = Rs("UsrSeq")
				End If
				
			End If
			
		End With
		
		Rs.Close
		Set Rs = Nothing

	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			Response.Write "SUCC"
		ElseIf Result = 1 Then
			Response.Write "EXISTS"
		Else 
			Response.Write "FAIL"
		End If
		Response.End

	End Sub

%>
