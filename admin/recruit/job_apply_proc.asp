<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_apply_proc.asp / 지원자 정보 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq
	Dim strParamFlag
	Dim Result
	Dim RecrSeq, Status
	Dim xKey, xWord, xStatus, xListSize

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		Flag 		= FN_Req("Flag","")
		Seq 		= FN_Req("Seq","0")

		RecrSeq 	= FN_Req("RecrSeq","0")
		Status 		= FN_Req("Status","")

		xKey 		= FN_Req("xKey","")
		xWord 		= FN_Req("xWord","")
		xStatus 	= FN_Req("xStatus","")
		xListSize 	= FN_Req("xListSize","")
		
		If Flag = "DEL" Or Flag = "STA" Then
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		Else
			strParamFlag = "@AdmSeq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
			'If AdmID = "" Or AdmLevel = "" Then 	Call SB_ReturnErr("[인수] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
	End Sub
	
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 	= Seq

			.Parameters("@Status").Value 			= Status
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		
		If Result = 0 Then
			strMsg = "처리가 완료됐습니다."
			
			If Flag = "ADD" Or Flag = "MOD" Then
				strScript = "opener.location.reload(); self.close();"
			Else
				strScript = "document.frmPage.action = ""job_apply_user.asp""; document.frmPage.submit(); "
			End If
		ElseIf Result = 1 Then
			strMsg = "이미 등록되어 있는 아이디 입니다."
			strScript = "history.back()"
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>지원자 관리 정보 처리 - HC 관리자</title>
		</head>
		<body>
			<form method="POST" name="frmPage" id="frmPage">
				<input type="hidden" name="RecrSeq" id="RecrSeq" value="<%=RecrSeq%>">
				<input type="hidden" name="sStatus" id="sStatus" value="<%=xStatus%>">
				<input type="hidden" name="sKey" id="sKey" value="<%=xKey%>">
				<input type="hidden" name="sWord" id="sWord" value="<%=xWord%>">
				<input type="hidden" name="ListSize" id="ListSize" value="<%=xListSize%>">
			</form>
			<script type="text/javascript">
			<!--
				alert("<%=strMsg%>");
				<%=strScript%>;
			//-->
			</script>

		</body>
	</html>
	<%
	End Sub
	
%>
