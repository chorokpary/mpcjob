<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pjt_proc.asp / 프로젝트 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-28 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq, sFlag, sView
	Dim PrjName, LogoPath, MapPath, Addr, Part, PrjFlag
	Dim TransSubway, TransBus, Location, Center, IsView
	Dim DelLogoPath, DelMapPath
	Dim LogoFName, MapFName
	Dim strParamFlag
	Dim Result
	
	Dim FilePath_Logo 	: FilePath_Logo = "cate"
	Dim FilePath_Map 	: FilePath_Map = "cate_map"
	Dim clsObjUpload : clsObjUpload = null

	Call Main()
	
	Sub Main()

		If FN_Req("Flag","") = "STA" Then
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

		sFlag 		= FN_Req("sFlag","")
		sView 		= FN_Req("sView","0")
	
		strParamFlag = "@StrSeq"
		Seq = Replace(Seq," ","")
		
	End Sub

	Sub getParamObj()

		With clsObjUpload
			Flag 		= FN_StrInj(.getReq("Flag"))
			Seq 		= FN_StrInj(.getReq("Seq"))
	
			PrjName 	= FN_StrInj(.getReq("PrjName"))
			Addr 		= FN_StrInj(.getReq("Addr"))
			Part 		= FN_StrInj(.getReq("Part"))
			PrjFlag 	= FN_StrInj(.getReq("PrjFlag"))
			TransSubway = FN_StrInj(.getReq("TransSubway"))
			TransBus 	= FN_StrInj(.getReq("TransBus"))
			Location 	= FN_StrInj(.getReq("Location"))
			Center 		= FN_StrInj(.getReq("Center"))
			
			DelLogoPath = FN_StrInj(.getReq("DelLogoPath"))
			DelMapPath 	= FN_StrInj(.getReq("DelMapPath"))
			
			// ### Attach File Upload : Start
			If .SaveFile(FilePath_Logo,"LogoPath") Then
				LogoFName = .getFileName()
				If LogoFName <> "" Then 
					LogoPath = GLOBAL_FILE_PATH & FilePath_Logo & "/" & LogoFName
				Else
					LogoPath = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If

			If .SaveFile(FilePath_Map,"MapPath") Then
				MapFName = .getFileName()
				If MapFName <> "" Then 
					MapPath = GLOBAL_FILE_PATH & FilePath_Map & "/" & MapFName
				Else
					MapPath = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Attach File Upload : End
			
		End With
	

		strParamFlag = "@Seq"
		If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		If PrjName = "" Then 	Call RtnErrorObj("[PARAM] 잘못된 경로로 접근하셨습니다.")

		
	End Sub
		
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspProject_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 	= Seq
			.Parameters("@PrjName").Value 		= PrjName
			.Parameters("@LogoPath").Value 		= LogoPath
			.Parameters("@MapPath").Value 		= MapPath
			.Parameters("@Addr").Value 			= Addr
			.Parameters("@Part").Value 			= Part
			.Parameters("@PrjFlag").Value 		= PrjFlag
			.Parameters("@TransSubway").Value 	= TransSubway
			.Parameters("@TransBus").Value 		= TransBus
			.Parameters("@Location").Value 		= Location
			.Parameters("@Center").Value 		= Center
			.Parameters("@DelLogoFlag").Value 	= Fn_SetDefault(DelLogoPath,"","0","1")
			.Parameters("@DelMapFlag").Value 	= Fn_SetDefault(DelMapPath,"","0","1")

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
			If Not FN_isBlank(DelLogoPath) Then  	Call clsObjUpload.Delete(strPath & "/" & DelLogoPath)
			If Not FN_isBlank(DelMapPath) Then  	Call clsObjUpload.Delete(strPath & "/" & DelMapPath)
			
			strMsg = "처리가 완료됐습니다."
			
			If Flag = "ADD" Or Flag = "MOD" Then
				strScript = "opener.location.reload(); self.close();"
			Else
				strScript = "window.location.href = ""pjt_list.asp?sFlag=" & sFlag & "&sView=" & sView & """; "
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
			<title>프로젝트 정보 처리 - HC 관리자</title>
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
