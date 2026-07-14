<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_proc.asp / 채용공고 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-21 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, FlagSub, Seq, DtSize, Token, IsAjax
	Dim Page, sViewFlag, sBDate, sEDate, sWord, ListSize, IsView
	
	Dim PrjSeq, Title, IsMain, IsTop, ChargeTask, TaskType
	Dim Part, RecrPart, RecrPart2, Career, Pay, HireType, LocalSido, LocalGugun, Times
	Dim AppBDate, AppEDate, LimitEdu, LimitCareer, LimitBAge, LimitEAge, LimitAgeNone, LimitGender
	Dim ChargeCode, ChargePhone, ChargeEmail, Contents, AdmSeq
	Dim MapPath, MapName
	
	Dim QueSeq, QueSort, Question, QueIsDisp

	Dim DelMapPath
	
	Dim strParamFlag
	Dim Result
	
	Dim FilePath 	: FilePath = "recruit"
	Dim clsObjUpload : clsObjUpload = null

	Call Main()
	
	Sub Main()
	
		Flag = FN_Req("Flag","")
		
		If Flag = "FIN" Or Flag = "NO_RECOM" Or Flag = "QUE"  Then	' STA : 마감  / NO_RECOM : 추천해제 / QUE : 사전질문 등록, 수정, 삭제
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
		
		If Flag = "QUE" Then
			Dim fnI
			For fnI = 0 To DtSize
				Call setQuestionProc(fnI+1)
			Next
		Else
			Call setProc()
		End If
		
		Call setResult()
		
	End Sub
	
	Sub getParamReq()

		Flag 	= FN_Req("Flag","")
		FlagSub	= FN_Req("FlagSub","")
		Seq 		= FN_Req("Seq","0")
		Token 	= FN_Req("Token","")
		IsAjax 	= FN_Req("IsAjax","0")
		
		sViewFlag = FN_Req("sViewFlag","")

		strParamFlag = "@StrSeq"
		Seq = Replace(Seq," ","")
		
		DtSize = FN_Req("DtSize","0")

		' ### Question : Start
		QueSeq = FN_Req("QueSeq","0") 
		QueSort = FN_Req("QueSort","1")
		Question = FN_Req("Question","")
		QueIsDisp = FN_Req("QueIsDisp","0")
		' ### Question : End
		
	End Sub

	Sub getParamObj()

		With clsObjUpload

			Flag = FN_StrInj(.getReq("Flag"))
			Seq 	= FN_StrInj(.getReq("Seq"))
			Token = FN_StrInj(.getReq("Token"))
			IsAjax = FN_StrInj(.getReq("IsAjax"))
			
			PrjSeq 	= FN_StrInj(.getReq("PrjSeq"))
			Title		= FN_StrInj(.getReq("Title"))
			IsMain 	= FN_StrInj(.getReq("IsMain"))
			IsTop 	= FN_StrInj(.getReq("IsTop"))
			ChargeTask = FN_StrInj(.getReq("ChargeTask"))
			TaskType = FN_StrInj(.getReq("TaskType"))
			Part 		= FN_StrInj(.getReq("Part"))
			RecrPart 	= FN_StrInj(.getReq("RecrPart"))
			RecrPart2 = FN_StrInj(.getReq("RecrPart2"))
			Career 	= FN_StrInj(.getReq("Career"))
			Pay 		= FN_StrInj(.getReq("Pay"))
			Times 		= FN_StrInj(.getReq("Times"))
			HireType 	= FN_StrInj(.getReq("HireType"))
			LocalSido = FN_StrInj(.getReq("LocalSido"))
			LocalGugun = FN_StrInj(.getReq("LocalGugun"))
			AppBDate = FN_StrInj(.getReq("AppBDate"))
			AppEDate = FN_StrInj(.getReq("AppEDate"))
			MapPath 	= FN_StrInj(.getReq("MapPath"))
			LimitEdu 	= FN_StrInj(.getReq("LimitEdu"))
			LimitCareer = FN_StrInj(.getReq("LimitCareer"))
			LimitBAge 	= FN_StrInj(.getReq("LimitBAge"))
			LimitEAge 	= FN_StrInj(.getReq("LimitEAge"))
			LimitAgeNone = FN_StrInj(.getReq("LimitAgeNone"))
			LimitGender = FN_StrInj(.getReq("LimitGender"))
			ChargeCode 	= FN_StrInj(.getReq("ChargeCode"))
			ChargePhone = FN_StrInj(.getReq("ChargePhone"))
			ChargeEmail = FN_StrInj(.getReq("ChargeEmail"))
			Contents 	= FN_StrInj(.getReq("Contents"))
			AdmSeq 		= FN_StrInj(.getReq("AdmSeq"))

			DelMapPath = FN_StrInj(.getReq("DelMapPath"))

			Page 		= FN_StrInj(.getReq("Page"))
			sViewFlag 	= FN_StrInj(.getReq("sViewFlag"))
			sBDate 		= FN_StrInj(.getReq("sBDate"))
			sEDate 		= FN_StrInj(.getReq("sEDate"))
			sWord 		= FN_StrInj(.getReq("sWord"))
			ListSize 	= FN_StrInj(.getReq("ListSize"))
			IsView 		= FN_StrInj(.getReq("IsView"))
			
			AdmSeq = Replace(AdmSeq," ","")

			If IsMain = "" Then			IsMain = 0
			If IsTop = "" Then			IsTop = 0
			If ChargeCode = "" Or Not IsNumeric(ChargeCode) Then 	ChargeCode = 0
			
			If LimitAgeNone = "1" Then
				LimitBAge = 0
				LimitEAge = 0
			End If
			
			// ### Img File Upload : Start
			If .SaveFile(FilePath,"MapPath") Then
				MapName = .getFileName()
				If MapName <> "" Then 
					MapPath = GLOBAL_FILE_PATH & FilePath & "/" & MapName
				Else
					MapPath = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Img File Upload : End
			
			If Flag = "REF" And MapPath = "" Then
				MapPath = FN_StrInj(.getReq("TmpImg"))
			End If
			
		End With


		strParamFlag = "@Seq"
		If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		If Title = "" Then 	Call RtnErrorObj("[PARAM] 잘못된 경로로 접근하셨습니다.")

		
	End Sub
		
	Sub setProc()
	
		Dim Rs

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		If StrComp("OK", returnMsg)= 0 Then

			With objCmd
				.ActiveConnection = objDbCon
				.CommandType = 4
				.CommandText ="uspRecruit_Proc"
				.Parameters.Refresh
				.Parameters("@Flag").Value 		= Flag
				.Parameters(strParamFlag).Value 	= Seq
				.Parameters("@Token").Value 	= Token
				
				.Parameters("@PrjSeq").Value 	= PrjSeq
				.Parameters("@Title").Value 		= Title
				.Parameters("@IsMain").Value 	= IsMain
				.Parameters("@IsTop").Value 	= IsTop
				.Parameters("@ChargeTask").Value = ChargeTask
				.Parameters("@TaskType").Value 	= TaskType
				.Parameters("@Part").Value 		= Part
				.Parameters("@RecrPart").Value 	= RecrPart
				.Parameters("@RecrPart2").Value 	= RecrPart2
				.Parameters("@Career").Value 	= Career
				.Parameters("@Pay").Value 		= Pay
				.Parameters("@HireType").Value 	= HireType
				.Parameters("@LocalSido").Value 	= LocalSido
				.Parameters("@LocalGugun").Value = LocalGugun
				.Parameters("@AppBDate").Value 	= AppBDate
				.Parameters("@AppEDate").Value 	= AppEDate
				.Parameters("@MapPath").Value 	= MapPath
				.Parameters("@LimitEdu").Value 	= LimitEdu
				.Parameters("@LimitCareer").Value = LimitCareer
				.Parameters("@LimitBAge").Value 	= LimitBAge
				.Parameters("@LimitEAge").Value 	= LimitEAge
				.Parameters("@LimitGender").Value = LimitGender
				.Parameters("@ChargeCode").Value = ChargeCode
				.Parameters("@ChargePhone").Value = oSeed.Encrypt(ChargePhone) 
				.Parameters("@ChargeEmail").Value = ChargeEmail
				.Parameters("@Contents").Value 	= Contents
				.Parameters("@AdmSeq").Value 	= AdmSeq
				
				.Parameters("@RecomFlag").Value = sViewFlag
				.Parameters("@DelImgFlag").Value = Fn_SetDefault(DelMapPath,"","0","1")
				.Parameters("@Times").Value 	= Times

				.Execute

				Result = .Parameters("@Result")
			
			End With
		

		Else
			Result = 1
		End If 
		
		Set oSeed = Nothing

	End Sub

	Sub setQuestionProc(argSeq)
		If FlagSub = "EDIT" Then
			QueSeq = FN_Req("QueSeq_"&argSeq,"0")
			QueSort = FN_Req("QueSort_"&argSeq,"1")
			Question = FN_Req("Question_"&argSeq,"")
			QueIsDisp = FN_Req("QueIsDisp_"&argSeq,"0")
		End If
				
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruitQaQ_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= FlagSub
			.Parameters("@RecrSeq").Value 	= Seq
			.Parameters("@RecrToken").Value = Token

			.Parameters("@QueSeq").Value	= QueSeq
			.Parameters("@Sort").Value 		= QueSort
			.Parameters("@Question").Value 	= Question
			.Parameters("@IsDisp").Value 	= QueIsDisp
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub
		
	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		Dim strPath : strPath = Server.MapPath("/")
		
		If IsAjax = "1" Then
			If Result = 0 Then
				Response.Write "SUCC"
			Else
				Response.Write "FAIL"
			End If
			Response.End
		Else
			If Result = 0 Then
				' ## 파일 삭제
				If Not FN_isBlank(DelMapPath) Then  	Call clsObjUpload.Delete(strPath & "/" & DelMapPath)
				
				strMsg = "처리가 완료됐습니다."
				
				If Flag = "ADD" OR Flag = "REF" Then
					strScript = "window.location.href = ""http://www.mpcjob.co.kr/admin/recruit/job_skin_list.asp""; "
				ElseIf Flag = "MOD" Then
					strScript = "$(""#frmPage"").attr({action:""http://www.mpcjob.co.kr/admin/recruit/job_skin_" & Fn_SetDefault(IsView,True,"list","elist") & ".asp"", method:""post""}).submit();"
				ElseIf Flag = "NO_RECOM" Then
					strScript = "window.location.href = ""http://www.mpcjob.co.kr/admin/recruit/job_recom_list.asp?sViewFlag=" & sViewFlag & """; "
				Else
					strScript = "window.location.href = ""http://www.mpcjob.co.kr/admin/recruit/job_skin_elist.asp""; "
				End If
			Else
				strMsg = "처리도중 오류가 발생했습니다."
				strScript = "history.back()"
			End If
		End If

		If Not IsNull(clsObjUpload) Then 	Set clsObjUpload = Nothing
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
			<title>채용 정보 처리 - HC 관리자</title>
			<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
			<script type="text/javascript">
			<!--
				$(window).load( function(){
					alert("<%=strMsg%>");
					<%=strScript%>
				} );
			//-->
			</script>
		</head>
		<body>
			<form name="frmPage" id="frmPage">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				<input type="hidden" id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
				<input type="hidden" id="sBDate" name="sBDate" value="<%=sBDate%>">
				<input type="hidden" id="sEDate" name="sEDate" value="<%=sEDate%>">
				<input type="hidden" id="sWord" name="sWord" value="<%=sWord%>">
				<input type="hidden" id="ListSize" name="ListSize" value="<%=ListSize%>">
				<input type="hidden" id="IsView" name="IsView" value="<%=IsView%>">
			</form>
		</body>
	</html>
	<%
	End Sub

	Sub RtnErrorObj (argMsg) 
		Set clsObjUpload = Nothing
		Call SB_ReturnErr(argMsg,"BACK")
	End Sub	
%>
