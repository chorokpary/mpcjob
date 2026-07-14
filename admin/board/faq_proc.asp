<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : faq_proc.asp / FAQ 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-02 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq, sKey, sWord, sCate
	Dim Subject, Passwd, Attach1_Path, Attach1_Name, Attach2_Path, Attach2_Name, Contents, Category
	Dim Attach1, Attach2
	Dim DelAttach1, DelAttach2
	Dim strParamFlag
	Dim Result
	
	Dim FilePath 	: FilePath = "board"
	Dim clsObjUpload : clsObjUpload = null

	Call Main()
	
	Sub Main()

		If FN_Req("Flag","") = "DEL" Then
			Call getParamReq()
		Else
			Set clsObjUpload = New ClsFileUpload

			' Create Object
			Dim DefPath : DefPath = Server.MapPath(GLOBAL_FILE_PATH)
			Dim IsObject : IsObject = clsObjUpload.setCreateObj(GLOBAL_UPLOAD_COMP, DefPath, GLOBAL_UPLOAD_SIZE)
			
			If Not IsObject Then 	Call RtnErrorObj("[OBJECT] 업로드 컴포넌트 설정 환경을 확인해 주세요.")

			' Request & File Upload
			Call getParamObj()
		End If

		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParamReq()

		Flag 		= FN_Req("Flag","")
		Seq 		= FN_Req("Seq","0")

		sCate 		= FN_Req("sCate","")
		sKey 		= FN_Req("sKey","")
		sWord 		= FN_InputVal(FN_Req("sWord",""))
	
		strParamFlag = "@StrSeq"
		Seq = Replace(Seq," ","")
		
	End Sub

	Sub getParamObj()

		With clsObjUpload
			Flag 		= FN_StrInj(.getReq("Flag"))
			Seq 		= FN_StrInj(.getReq("Seq"))
	
			Subject 	= FN_StrInj(.getReq("Subject"))
			Category 	= FN_StrInj(.getReq("Category"))
			Passwd 		= FN_StrInj(.getReq("Passwd"))
			Contents 	= FN_StrInj(.getReq("Contents"))
			
			DelAttach1 	= FN_StrInj(.getReq("DelAttach1"))
			DelAttach2 	= FN_StrInj(.getReq("DelAttach2"))
			
			// ### Attach File Upload : Start
			If .SaveFile(FilePath,"Attach1") Then
				Attach1_Name = .getFileName()
				If Attach1_Name <> "" Then 
					Attach1_Path = GLOBAL_FILE_PATH & FilePath & "/" & Attach1_Name
				Else
					Attach1_Path = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If

			If .SaveFile(FilePath,"Attach2") Then
				Attach2_Name = .getFileName()
				If Attach2_Name <> "" Then 
					Attach2_Path = GLOBAL_FILE_PATH & FilePath & "/" & Attach2_Name
				Else
					Attach2_Path = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Attach File Upload : End
			
		End With

		strParamFlag = "@Seq"
		If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		If Subject = "" Then 	Call RtnErrorObj("[PARAM] 잘못된 경로로 접근하셨습니다.")

	End Sub
		
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspBbs_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 	= Seq
			.Parameters("@BbsKey").Value 		= "faq"
			.Parameters("@Subject").Value 		= Subject
			.Parameters("@Category").Value 		= Category
			.Parameters("@Contents").Value 		= Contents
			.Parameters("@Attach1_Path").Value 	= Attach1_Path
			.Parameters("@Attach1_Name").Value 	= Attach1_Name
			.Parameters("@Attach2_Path").Value 	= Attach2_Path
			.Parameters("@Attach2_Name").Value 	= Attach2_Name
			.Parameters("@IsAdmin").Value 		= 1
			.Parameters("@RegIP").Value 		= Request.ServerVariables("REMOTE_ADDR")
			.Parameters("@RegSeq").Value 		= Session("ASeq")
			.Parameters("@RegName").Value 		= Session("AName")
			.Parameters("@Passwd").Value 		= Passwd
			.Parameters("@DelAttach1").Value 	= Fn_SetDefault(DelAttach1,"","0","1")
			.Parameters("@DelAttach2").Value 	= Fn_SetDefault(DelAttach2,"","0","1")

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
			If Not FN_isBlank(DelAttach1) Then  	Call clsObjUpload.Delete(strPath & "/" & DelAttach1)
			If Not FN_isBlank(DelAttach2) Then  	Call clsObjUpload.Delete(strPath & "/" & DelAttach2)
			
			strMsg = "처리가 완료됐습니다."
			
			If Flag = "ADD" Or Flag = "MOD" Then
				strScript = "opener.location.reload(); self.close();"
			Else
				strScript = "window.location.href = ""faq_list.asp?sCate=" & sCate & """; "
			End If
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If

		If Not IsNull(clsObjUpload) Then 	Set clsObjUpload = Nothing
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>FAQ 정보 처리 - HC 관리자</title>
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
