<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pop_proc.asp / 팝업정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-27 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq, Cate, sView
	Dim Title, IsView, Width, Height, BDate, EDate, Contents
	Dim strParamFlag
	Dim Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		Flag 		= FN_Req("Flag","")
		Seq 		= FN_Req("Seq","0")
		sView 		= FN_Req("sView","1")
	
		Title 	= FN_Req("Title","")
		IsView 	= FN_Req("IsView","0")
		Width 	= FN_Req("Width","0")
		Height 	= FN_Req("Height","0")
		BDate 	= FN_Req("BDate",Date())
		EDate 	= FN_Req("EDate",Date())
		Contents = FN_Req("Contents","")
		
		If Flag = "DEL" Then
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		Else
			strParamFlag = "@Seq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
			If Title = "" Or Not IsNumeric(Width) Or Not IsNumeric(Height) Then 	Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
	End Sub
	
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfPopup_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters(strParamFlag).Value = Seq
			.Parameters("@Title").Value 	= Title
			.Parameters("@IsView").Value 	= IsView
			.Parameters("@Width").Value 	= Width
			.Parameters("@Height").Value 	= Height
			.Parameters("@BDate").Value 	= BDate
			.Parameters("@EDate").Value 	= EDate
			.Parameters("@Contents").Value 	= Contents

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			strMsg = "처리가 완료됐습니다."
			strScript = "window.location.href = ""pop_list.asp""; "
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
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
