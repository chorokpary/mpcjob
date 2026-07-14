<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : message_proc.asp / 쪽지 답변, 삭제 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-28 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq
	Dim Tel, RecvSeq, Answer
	Dim strParamFlag
	Dim Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		Flag 	= FN_Req("Flag","")
		Seq 	= FN_Req("Seq","0")
		
		Tel 	= FN_Req("Tel","")
		RecvSeq = FN_Req("RecvSeq","")
		Answer 	= FN_Req("Answer","")
		
		
		Answer 	= FN_ClearTag(Answer)
		
		If Flag = "DEL_RECV" Then
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		Else
			strParamFlag = "@Seq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
			If Tel = "" Or Not IsNumeric(Replace(Tel,"-","")) Or Answer = "" Then 	Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
	End Sub
	
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspNote_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters(strParamFlag).Value = Seq

			.Execute
			
			Result = .Parameters("@Result")
		End With
		
		If Flag = "ANS" Then		Call SB_SendSMS(Tel,RecvSeq,"L","[채용문의답변]",Answer,0,"","NOTE",Seq,"")
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			strMsg = "처리가 완료됐습니다."
			strScript = "window.location.href = ""message_list.asp""; "
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>쪽지 정보 처리 - MPC 관리자</title>
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
