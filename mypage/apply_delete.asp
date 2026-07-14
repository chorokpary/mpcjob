<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : apply_delete.asp / HC - 입사지원 취소
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim Seq, Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub

	Sub getParam()
	
		Seq 	= FN_Req("Seq","")
		
	End Sub
	
	Sub setProc()
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_Proc"
			.Parameters.Refresh
			'.Parameters("@Flag").Value 		= "DEL"
			.Parameters("@Flag").Value 		= "CANCEL"
			.Parameters("@StrSeq").Value 	= Seq

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			' ## 파일 삭제
			strMsg = "처리가 완료됐습니다."
			strScript = "window.location.href = ""apply.asp""; "
		Else
			strMsg = "[" & Result & "] 처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>입사지원 취소 처리 - HC 관리자</title>
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
