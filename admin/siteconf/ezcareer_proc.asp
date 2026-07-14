<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : ezcareer_proc.asp / 경력코드 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2013-07-24 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%

	Dim Flag, Sort, ChSortCur, CodeLen
	Dim CONF_TBL, CONF_FLD, Key, IsUse
	Dim strParamFlag
	Dim Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		CONF_TBL = "TBL_USER_CAREER_EZ"
		CONF_FLD = "Career"

		Flag 		= FN_Req("Flag","")
		Sort 		= FN_Req("Sort","")
		ChSortCur 	= FN_Req("ChSortCur","")
		CodeLen 	= FN_Req("CodeLen","0")
	
		Key 		= FN_Req("Key","")
		IsUse 		= FN_Req("IsUse","0")
		
		Key = Replace(Key,",","&#44;")	'checkbox용 코드이기 때문에 콤마 사용 불가

		If Flag = "STA" Then
			strParamFlag = "@StrSort"
			Sort = Replace(Sort," ","")
		ElseIf Flag = "SORT" Then
			strParamFlag = "@Sort"
			Sort = FN_Req("ChSort","0")
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0
		Else
			strParamFlag = "@Sort"
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0
		End If
		
		If Not IsNumeric(IsUse) Then 	IsUse = 0
		
	End Sub
	
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfCode_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters(strParamFlag).Value = Sort
			.Parameters("@Tbl").Value 		= CONF_TBL
			.Parameters("@Field").Value 	= CONF_FLD
			.Parameters("@CdKey").Value 	= Key
			.Parameters("@IsUse").Value 	= IsUse
			.Parameters("@SortCur").Value 	= ChSortCur

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			If Flag <> "SORT" Then
				strMsg = "처리가 완료됐습니다."
			End If
			strScript = "window.location.href = ""ezcareer_list.asp"";"
		ElseIf Result = 1 Then
			If Flag <> "SORT" Then
				strMsg = "이미 등록되어 있습니다."
			End If
			strScript = "window.location.href = ""ezcareer_list.asp"";"
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"/>
			<title>코드 정보 처리 - HC 관리자</title>
			<script type="text/javascript">
			<!--
				<% If strMsg <> "" Then %>
				alert("<%=strMsg%>");
				<% End If %>
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
