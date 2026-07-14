<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : article_proc.asp / HC - 정회원 전환 처리 (Ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim Result

	Call Main()
	
	Sub Main()
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
			.Parameters("@Flag").Value 		= "UPGRADE"
			.Parameters("@Seq").Value 		= Session("USeq")

			Set Rs = .Execute
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = 99
			End If
			
		End With
		
		Set Rs = Nothing

		If Result = "0" Then
			Session("UType") = 1
		End If
		
		Set Rs = Nothing
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			%><meta http-equiv="refresh" content="0; url=/mypage/resume.asp"><%
		Else 
			%>
			<script type="text/javascript">
				alert("정회원 전환중 오류가 발생했습니다. 다시 시도하세요.");
				history.back();
			</script>
			<%
		End If
		Response.End

	End Sub
%>
