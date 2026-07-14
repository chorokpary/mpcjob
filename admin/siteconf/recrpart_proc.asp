<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : recrpart_proc.asp / 직종 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2013-03-15 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%

	Dim Flag, ChSortCur, CodeLen
	Dim CONF_TBL, CONF_FLD, CONF_TBL2, CONF_FLD2
	DIm Sort, Sort2, Key, Key2, IsUse, ChangeAll
	Dim strParamFlag
	Dim Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		CONF_TBL = "TBL_RECRUIT"
		CONF_FLD = "RecrPart"

		CONF_TBL2 = "TBL_RECRUIT"
		CONF_FLD2 = "RecrPart2"

		Flag 		= FN_Req("Flag","")
		Sort 		= FN_Req("Sort","")
		Sort2 		= FN_Req("Sort2","0")
		ChSortCur 	= FN_Req("ChSortCur","")
		CodeLen 	= FN_Req("CodeLen","0")
	
		Key 		= FN_Req("Key","")
		Key2 		= FN_Req("Key2","")
		IsUse 		= FN_Req("IsUse","0")
		
		ChangeAll 	= FN_Req("ChangeAll","N")

		If Flag = "STA" Then
			strParamFlag = "@StrSort"
			Sort = Replace(Sort," ","")
		ElseIf Flag = "SORT" Then
			strParamFlag = "@Sort"
			Sort = FN_Req("ChSort","0")
			Sort2 = FN_Req("ChSort2","0")
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0
			If Sort2 = "" Or Not IsNumeric(Sort2) Then 	Sort2 = 0
		Else
			strParamFlag = "@Sort"
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0
			If Sort2 = "" Or Not IsNumeric(Sort2) Then 	Sort2 = 0
		End If
		
		If Not IsNumeric(IsUse) Then 	IsUse = 0
		
	End Sub
	
	Sub setProc()
		With objCmd
		
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfCode_Proc_2Dep"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters(strParamFlag).Value = Sort

			.Parameters("@Tbl").Value 		= CONF_TBL
			.Parameters("@Field").Value 	= CONF_FLD
			.Parameters("@CdKey").Value 	= Key

			.Parameters("@Sort2").Value 	= Sort2
			.Parameters("@Tbl2").Value 		= CONF_TBL2
			.Parameters("@Field2").Value 	= CONF_FLD2
			.Parameters("@CdKey2").Value 	= Key2
			.Parameters("@IsUse").Value 	= IsUse
			
			.Parameters("@SortCur").Value 	= ChSortCur
			.Parameters("@ChangeAll").Value = ChangeAll

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
			strScript = "window.location.href = ""recrpart_list.asp"";"
		ElseIf Result = 1 Then
			If Flag <> "SORT" Then
				strMsg = "이미 등록되어 있습니다."
			End If
			strScript = "window.location.href = ""recrpart_list.asp"";"
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
