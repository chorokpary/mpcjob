<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_flash_proc.asp / 플래시 설정 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, sViewFlag, sPage, Result
	Dim Seq, SeqList, ImgPath, LinkUrl, Sort, BDate, EDate
	Dim ChSortCur
	Dim ImgName, DelImgPath

	Dim FilePath 	: FilePath = "main_banner"
	Dim clsObjUpload : clsObjUpload = null

	Dim strParamFlag

	Call Main()
	
	Sub Main()

		If FN_Req("Flag","") = "DEL" Or FN_Req("Flag","") = "SORT" Then
			Call getParamReq()
		Else
			Set clsObjUpload = New ClsFileUpload
	
			' Create Object
			Dim DefPath : DefPath = Server.MapPath(GLOBAL_FILE_PATH)
			Dim IsObject : IsObject = clsObjUpload.setCreateObj(GLOBAL_UPLOAD_COMP, DefPath, GLOBAL_UPLOAD_SIZE)
				
			If Not IsObject Then 	Call RtnErrorObj("[OBJECT] 업로드 컴포넌트 설정 환경을 확인해 주세요.")
	
			' Request & File Upload
			Call getParamObj()
			Call setProc()
		End If

		Call setResult()
	End Sub

	Sub getParamReq()

		Dim i, tmpSeq
		
		Flag 		= FN_Req("Flag","")
		sViewFlag 	= FN_Req("sViewFlag","")
		sPage 		= FN_Req("sPage","1")
		Seq 			= FN_Req("Seq","0")
		SeqList 		= FN_Req("SeqList","")
		
		If Flag = "SORT" Then
			strParamFlag = "@Seq"

			'For i = 1 To 4
			'	tmpSeq = FN_Req("Sort" & i,"0:0")
				
			'	Seq = Left(tmpSeq, InStr(tmpSeq,":")-1)
			'	Sort = Right(tmpSeq,1)
				
			'	Call setProc()
			'Next

			strParamFlag = "@Sort"
			Sort = FN_Req("ChSort","0")
			ChSortCur 	= FN_Req("ChSortCur","")
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0
				
			Call setProc()
		Else
			strParamFlag = "@StrSeq"
			Seq = Replace(SeqList," ","")
			
			Call setProc()
		End If

	End Sub
	
	Sub getParamObj()

		strParamFlag = "@Seq"
		
		With clsObjUpload
			Flag = FN_StrInj(.getReq("Flag"))
			sViewFlag = FN_StrInj(.getReq("sViewFlag"))
			Seq = FN_StrInj(.getReq("Seq"))
			LinkUrl = FN_StrInj(.getReq("LinkUrl"))
			Sort = FN_StrInj(.getReq("Sort"))
			BDate = FN_StrInj(.getReq("BDate"))
			EDate = FN_StrInj(.getReq("EDate"))
			
			If Sort = "" Or Not IsNumeric(Sort) Then 	Sort = 0


			' 삭제데이터
			DelImgPath = FN_StrInj(.getReq("DelImgPath"))
			
			// ### Attach File Upload : Start
			If .SaveFile(FilePath,"ImgPath") Then
				ImgName = .getFileName()
				If ImgName <> "" Then 
					ImgPath = GLOBAL_FILE_PATH & FilePath & "/" & ImgName
				Else
					ImgPath = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Attach File Upload : End

			If Flag = "REF" And ImgPath = "" Then
				ImgPath = FN_StrInj(.getReq("TmpImg"))
			End If
			
		End With
		
	End Sub
		
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfFlash_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters("@ViewFlag").Value 	= sViewFlag
			.Parameters(strParamFlag).Value = Seq

			.Parameters("@ImgPath").Value 	= ImgPath
			.Parameters("@LinkUrl").Value 	= LinkUrl
			.Parameters("@Sort").Value 		= Sort
			.Parameters("@BDate").Value 	= BDate
			.Parameters("@EDate").Value 	= EDate
			
			.Parameters("@DelImgPath").Value = Fn_SetDefault(DelImgPath,"","0","1")
			.Parameters("@SortCur").Value 	= ChSortCur

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub

	Function getUseFileCount()
		Dim Rs
		Dim fnResult : fnResult = 0
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfFlash_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FILE"
			.Parameters("@ImgPath").Value 	= DelImgPath
			
			' ## 이미지정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				fnResult = Rs("CntFile")
			End If
			
			Rs.Close
			Set Rs = Nothing

		End With
		
		getUseFileCount = fnResult
	End Function

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		Dim strPath : strPath = Server.MapPath("/")
		
		If Result = 0 Then
			' ## 파일 삭제
			If Not FN_isBlank(DelImgPath) And getUseFileCount() <= 0 Then  	Call clsObjUpload.Delete(strPath & "/" & DelImgPath)
			
			strScript = "window.location.href = ""conf_flash.asp?sViewFlag=" & sViewFlag 
			
			If Flag = "SORT" Then
				strMsg = ""
				strScript = strScript & "&amp;Page=" & sPage
			Else
				strMsg = "처리가 완료됐습니다."
			End If
			
			strScript = strScript & """; "
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
			<title>플래시설정 정보 처리 - HC 관리자</title>
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

	Sub RtnErrorObj (argMsg) 
		Set clsObjUpload = Nothing
		Call SB_ReturnErr(argMsg,"BACK")
	End Sub	
%>
