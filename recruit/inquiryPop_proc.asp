<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : inquiryPop_proc.asp / 채용문의 처리 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim RecrSeq
	Dim ChargeCode, ChargeName, SendName, Tel, Memo
	
	Dim Result
	
	Call Main()
	
	Sub Main()
	
		RecrSeq 	= FN_Req("RecrSeq","0")
		ChargeCode 	= FN_Req("ChargeCode","0")
		ChargeName 	= FN_Req("ChargeName","")
		SendName 	= FN_Req("SendName","")
		Tel 		= FN_Req("Tel_1","") & "-" & FN_Req("Tel_2","") & "-" & FN_Req("Tel_3","")
		Memo 		= FN_Req("Memo","")

		Call setProc()
		Call setResult()
		
	End Sub

	Sub setProc()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspNote_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADD"
			.Parameters("@Seq").Value 		= 0
			.Parameters("@RecrSeq").Value 	= RecrSeq
			
			.Parameters("@RecvSeq").Value 	= ChargeCode
			.Parameters("@RecvName").Value 	= ChargeName
			.Parameters("@RecvType").Value 	= "2"
			
			.Parameters("@SendSeq").Value 	= Session("USeq")
			.Parameters("@SendName").Value 	= Session("UName")
			.Parameters("@SendType").Value 	= "1"

			.Parameters("@Title").Value = Session("UName") & "님이 채용문의를 남기셨습니다."
			.Parameters("@Tel").Value 	= Tel
			.Parameters("@Memo").Value 	= Memo
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			Response.Write "SUCC"
		Else 
			Response.Write "FAIL"
		End If
		Response.End

	End Sub

%>
