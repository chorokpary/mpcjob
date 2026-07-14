<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_main_proc.asp / 메일 설정 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-22 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Result
	Dim ConfGreeting, ConfFormDoc, ConfFormXls, ConfFormHwp
	Dim ConfFormDoc_Name, ConfFormXls_Name, ConfFormHwp_Name
	Dim DelConfFormDoc, DelConfFormXls, DelConfFormHwp

	Dim FilePath 	: FilePath = "main_conf"
	Dim clsObjUpload : clsObjUpload = null

	Call Main()
	
	Sub Main()

		Set clsObjUpload = New ClsFileUpload

		' Create Object
		Dim DefPath : DefPath = Server.MapPath(GLOBAL_FILE_PATH)
		Dim IsObject : IsObject = clsObjUpload.setCreateObj(GLOBAL_UPLOAD_COMP, DefPath, GLOBAL_UPLOAD_SIZE)
			
		If Not IsObject Then 	Call RtnErrorObj("[OBJECT] 업로드 컴포넌트 설정 환경을 확인해 주세요.")

		' Request & File Upload
		Call getParamObj()

		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParamObj()

		With clsObjUpload
			ConfGreeting = FN_StrInj(.getReq("ConfGreeting"))
			ConfFormDoc = FN_StrInj(.getReq("ConfFormDoc"))
			ConfFormXls = FN_StrInj(.getReq("ConfFormXls"))
			ConfFormHwp = FN_StrInj(.getReq("ConfFormHwp"))

			' 삭제데이터
			DelConfFormDoc = FN_StrInj(.getReq("DelConfFormDoc"))
			DelConfFormXls = FN_StrInj(.getReq("DelConfFormXls"))
			DelConfFormHwp = FN_StrInj(.getReq("DelConfFormHwp"))
			
			// ### Attach File Upload : Start
			If .SaveFile(FilePath,"ConfFormDoc") Then
				ConfFormDoc_Name = .getFileName()
				If ConfFormDoc_Name <> "" Then 
					ConfFormDoc = GLOBAL_FILE_PATH & FilePath & "/" & ConfFormDoc_Name
				Else
					ConfFormDoc = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If

			If .SaveFile(FilePath,"ConfFormXls") Then
				ConfFormXls_Name = .getFileName()
				If ConfFormXls_Name <> "" Then 
					ConfFormXls = GLOBAL_FILE_PATH & FilePath & "/" & ConfFormXls_Name
				Else
					ConfFormXls = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If

			If .SaveFile(FilePath,"ConfFormHwp") Then
				ConfFormHwp_Name = .getFileName()
				If ConfFormHwp_Name <> "" Then 
					ConfFormHwp = GLOBAL_FILE_PATH & FilePath & "/" & ConfFormHwp_Name
				Else
					ConfFormHwp = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Attach File Upload : End
			
		End With
		
	End Sub
		
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfSite_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "DEF"
			.Parameters("@ConfGreeting").Value 	= ConfGreeting
			.Parameters("@ConfFormDoc").Value 	= ConfFormDoc
			.Parameters("@ConfFormXls").Value 	= ConfFormXls
			.Parameters("@ConfFormHwp").Value 	= ConfFormHwp
			
			.Parameters("@DelConfFormDoc").Value 	= Fn_SetDefault(DelConfFormDoc,"","0","1")
			.Parameters("@DelConfFormXls").Value 	= Fn_SetDefault(DelConfFormXls,"","0","1")
			.Parameters("@DelConfFormHwp").Value 	= Fn_SetDefault(DelConfFormHwp,"","0","1")

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		Dim strPath : strPath = Server.MapPath("/")
		
		If Result = 0 Then
			' ## 파일 삭제
			If Not FN_isBlank(DelConfFormDoc) Then  	Call clsObjUpload.Delete(strPath & "/" & DelConfFormDoc)
			If Not FN_isBlank(DelConfFormXls) Then  	Call clsObjUpload.Delete(strPath & "/" & DelConfFormXls)
			If Not FN_isBlank(DelConfFormHwp) Then  	Call clsObjUpload.Delete(strPath & "/" & DelConfFormHwp)
			
			strMsg = "처리가 완료됐습니다."
			strScript = "window.location.href = ""conf_main.asp""; "
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If

		If Not IsNull(clsObjUpload) Then 	Set clsObjUpload = Nothing
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"/>
			<title>메인설정 정보 처리 - HC 관리자</title>
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

	Sub RtnErrorObj (argMsg) 
		Set clsObjUpload = Nothing
		Call SB_ReturnErr(argMsg,"BACK")
	End Sub	
%>
