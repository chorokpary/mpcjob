<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : partner_proc.asp / 파트너(입력업체, 온라인광고)정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-22 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq, Cate, sView
	Dim PartnerName, BizNum, ChargeName, ChargeTel
	Dim strParamFlag
	Dim Result, H_SSL

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		Flag 		= FN_Req("Flag","")
		Seq 		= FN_Req("Seq","0")
		Cate 		= FN_Req("Cate","0")
		sView 		= FN_Req("sView","1")
	
		PartnerName = FN_Req("PartnerName","")
		BizNum 		= FN_Req("BizNum","")
		ChargeName 	= FN_Req("ChargeName","")
		ChargeTel 	= FN_Req("ChargeTel","")

		H_SSL 		= FN_Req("H_SSL","")

		If UCase(H_SSL) = "Y" Then  
		%>
		<script type="text/javascript">
			opener.location.reload(); self.close();	
		</script>
		<%
		Response.End
		End If 
		
		If Flag = "STA" Then
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		Else
			strParamFlag = "@Seq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
			If PartnerName = "" Or Cate = "0" Then 	Call SB_ReturnErr("[인수] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
	End Sub
	
	Sub setProc()

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		If StrComp("OK", returnMsg)= 0 Then

			With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspPartner_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 	= Seq
			.Parameters("@PartnerName").Value 	= PartnerName
			.Parameters("@BizNum").Value 		= BizNum
			.Parameters("@ChargeName").Value 	= ChargeName
			.Parameters("@ChargeTel").Value 	= oSeed.Encrypt(ChargeTel)
			.Parameters("@Parent").Value 		= Cate
			.Execute
			
			Result = .Parameters("@Result")
			
			End With

		Else
			Result = 1
		End If 
		
	
		Set oSeed = Nothing
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			strMsg = "처리가 완료됐습니다."
			
			If Flag = "ADD" Or Flag = "MOD" Then
				strScript = "window.location.href = ""http://www.mpcjob.co.kr/admin/siteconf/partner_proc.asp?H_SSL=Y""; "
				'strScript = "opener.location.reload(); self.close();"
			Else
				Select Case Cate
					Case "5" : strUrl = "partner_skin_olist.asp"
					Case "6" : strUrl = "partner_skin_clist.asp"
				End Select
			
				strScript = "window.location.href = """ & strUrl & "?sView=" & sView & """; "
			End If
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"/>
			<title>협력업체 정보 처리 - HC 관리자</title>
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
