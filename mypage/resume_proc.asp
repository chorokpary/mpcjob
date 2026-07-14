<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : resume_proc.asp / 이력서 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4m / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, FlagSub, Seq, sFlag, sView, IsAjax
	Dim PartnerDepth1, PartnerDepth2, RecrSeq
	Dim UsrName, Birthday, Gender, Mobile, ChkMobile, HopePart, HopePart2, AddrSido, AddrGugun, Addr, AddrDetail
	Dim Jumin1, Jumin2, EzCareer
	Dim Email, Zipcode1, Zipcode2, Intro, EduLast, EduType
	Dim Passwd
	Dim ImgPath, ImgName
	Dim Attach1_Path, Attach2_Path, Attach3_Path
	Dim Attach1_Name, Attach2_Name, Attach3_Name
	
	Dim EduBYear, EduBMonth, EduEYear, EduEMonth, SchoolName, Major, Status, Local
	Dim CrBYear, CrBMonth, CrEYear, CrEMonth, Reason, CompanyName, Department, Task, Position, CompanyYear
	Dim Relation, FmName, Age, Education, Job, IsLiveWith 
	Dim Memo
	Dim Sort
	
	Dim DelImgPath, DelAttach_1, DelAttach_2, DelAttach_3
	
	Dim strParamFlag
	Dim Result
	
	Dim FilePath_Att 	: FilePath_Att = "profile"
	Dim FilePath_Photo 	: FilePath_Photo = "photo"
	Dim clsObjUpload : clsObjUpload = null

	Call Main()
	
	Sub Main()
		
		Flag = FN_Req("Flag","")
		
		If Flag = "DEL" Or Flag = "BLK" Or Flag = "RECR" Or Flag = "PWD_C" Or Flag = "PWD_S" _
			Or Flag = "EDU" Or Flag = "CAREER" Or Flag = "FAM" Or Flag = "MEMO" Then
			' DEL : 관리자삭제 / BLK : 블랙리스트 처리 / RECR : 채용공고 지원 / PWD_C : 비밀번호변경 / PWD_S : 임시비밀번호 전송
			' EDU : 학력정보 / CAREER : 경력정보 / FAM : 가족사항 / MEMO : 관리자 메모
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


		If Flag = "RECR" Then
			Call setRecrProc()
		ElseIf Flag = "EDU" Then
			Call setEduProc()
		ElseIf Flag = "CAREER" Then
			Call setCareerProc()
		ElseIf Flag = "FAM" Then
			Call setFamProc()
		ElseIf Flag = "MEMO" Then
			Call setMemoProc()
		Else
			Call setProc()
		End If
		
		Call setResult()
	End Sub
	
	Sub getParamReq()

		Seq 		= Session("USeq")
		
		Flag 		= FN_Req("Flag","")
		FlagSub		= FN_Req("FlagSub","")
		IsAjax 		= FN_Req("IsAjax","0")

		sFlag 		= FN_Req("sFlag","")
		sView 		= FN_Req("sView","0")

		PartnerDepth1 	= FN_Req("PartnerDepth1","")
		PartnerDepth2 	= FN_Req("PartnerDepth2","")
		RecrSeq 		= FN_Req("RecrSeq","0")

		Passwd 		= FN_Req("Passwd","")
		
		If Passwd <> "" Then		Passwd = hex_sha1(Passwd)
		
		If Flag = "PWD_C" Or Flag = "PWD_S" Then
			strParamFlag = "@Seq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		Else	' 복수 변경이 가능한 Flag들..
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		End If

		' ### Education : Start
		EduBYear 	= FN_Req("EduBYear","")
		EduBMonth 	= FN_Req("EduBMonth","")
		EduEYear 	= FN_Req("EduEYear","")
		EduEMonth 	= FN_Req("EduEMonth","")
		SchoolName 	= FN_Req("SchoolName","")
		Major 		= FN_Req("Major","")
		Status 		= FN_Req("Status","")
		Local 		= FN_Req("Local","")
		' ### Education : End
		
		' ### Career : Start
		CrBYear 	= FN_Req("CrBYear","")
		CrBMonth 	= FN_Req("CrBMonth","")
		CrEYear 	= FN_Req("CrEYear","")
		CrEMonth 	= FN_Req("CrEMonth","")
		Reason 		= FN_Req("Reason","")
		CompanyName = FN_Req("CompanyName","")
		Department 	= FN_Req("Department","")
		Task 		= FN_Req("Task","")
		Position 	= FN_Req("Position","")
		CompanyYear = FN_Req("CompanyYear","")
		' ### Career : End
		
		' ### Familly : Start
		Relation 	= FN_Req("Relation","")
		FmName 		= FN_Req("FmName","")
		Age 		= FN_Req("Age","")
		Education 	= FN_Req("Education","")
		Job 		= FN_Req("Job","")
		IsLiveWith 	= FN_Req("IsLiveWith","")

		If Age = "" Or Not IsNumeric(Age) Then 					Age = Null
		If IsLiveWith = "" Or Not IsNumeric(IsLiveWith) Then 	IsLiveWith = Null
		' ### Familly : End

		' ### Memo : Start
		Memo 	= FN_Req("Memo","")
		' ### Memo : End
		
		Sort 	= FN_Req("Sort","")
		Sort = Replace(Sort," ","")
		
	End Sub

	Sub getParamObj()

		With clsObjUpload
			Seq 		= Session("USeq")
			
			Flag 		= FN_StrInj(.getReq("Flag"))
			IsAjax 		= FN_StrInj(.getReq("IsAjax"))
			
			UsrName 	= FN_StrInj(.getReq("UsrName"))
			
			If FN_StrInj(.getReq("Birthday_Y")) <> "" Then
				Birthday 	= FN_StrInj(.getReq("Birthday_Y")) & "-" & FN_StrInj(.getReq("Birthday_M")) & "-" & FN_StrInj(.getReq("Birthday_D"))
			Else
				Birthday = ""
			End If
			
			Gender 	= FN_StrInj(.getReq("Gender"))
			Mobile 	= FN_StrInj(.getReq("Mobile"))
			ChkMobile = FN_StrInj(.getReq("ChkMobile"))
			HopePart 	= FN_StrInj(.getReq("HopePart"))
			HopePart2 = FN_StrInj(.getReq("HopePart2"))
			AddrSido 	= FN_StrInj(.getReq("AddrSido"))
			AddrGugun = FN_StrInj(.getReq("AddrGugun"))
			Addr 	= FN_StrInj(.getReq("Addr"))
			AddrDetail = FN_StrInj(.getReq("AddrDetail"))
			Email 	= FN_StrInj(.getReq("Email"))
			Zipcode1 = FN_StrInj(.getReq("Zipcode1"))
			Zipcode2 = FN_StrInj(.getReq("Zipcode2"))

			Jumin1 	= FN_StrInj(.getReq("Jumin1"))
			Jumin2 	= FN_StrInj(.getReq("Jumin2"))
			EzCareer = FN_StrInj(.getReq("EzCareer"))
			
			Intro  	= FN_StrInj(.getReq("Intro"))

			EduLast  	= FN_StrInj(.getReq("EduLast"))
			EduType  	= FN_StrInj(.getReq("EduType"))
			
			DelImgPath = FN_StrInj(.getReq("DelImgPath"))
			DelAttach_1 = FN_StrInj(.getReq("DelAttach_1"))
			DelAttach_2 = FN_StrInj(.getReq("DelAttach_2"))
			DelAttach_3 = FN_StrInj(.getReq("DelAttach_3"))
			
			If  Addr = "" Then
				Addr = Trim(AddrSido & " " & AddrGugun)
			End If
			
			// ### Img File Upload : Start
			If .SaveFile(FilePath_Photo,"ImgPath") Then
				ImgName = .getFileName()
				If ImgName <> "" Then 
					ImgPath = GLOBAL_FILE_PATH & FilePath_Photo & "/" & ImgName
				Else
					ImgPath = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Img File Upload : End
			
			// ### Attach File Upload : Start
			If .SaveFile(FilePath_Att,"Attach_1") Then
				Attach1_Name = .getFileName()
				If Attach1_Name <> "" Then 
					Attach1_Path = GLOBAL_FILE_PATH & FilePath_Att & "/" & Attach1_Name
				Else
					Attach1_Path = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			
			If .SaveFile(FilePath_Att,"Attach_2") Then
				Attach2_Name = .getFileName()
				If Attach2_Name <> "" Then 
					Attach2_Path = GLOBAL_FILE_PATH & FilePath_Att & "/" & Attach2_Name
				Else
					Attach2_Path = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If

			If .SaveFile(FilePath_Att,"Attach_3") Then
				Attach3_Name = .getFileName()
				If Attach3_Name <> "" Then 
					Attach3_Path = GLOBAL_FILE_PATH & FilePath_Att & "/" & Attach3_Name
				Else
					Attach3_Path = ""
				End If
			Else
				Call RtnErrorObj(.getErrMsg)
			End If
			// ### Attach File Upload : End
			
		End With
	

		strParamFlag = "@Seq"
		If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0

		
	End Sub
		
	Sub setProc()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 	= Seq
			
			.Parameters("@Passwd").Value 		= Passwd
			.Parameters("@UsrName").Value 		= UsrName
			.Parameters("@Birthday").Value 		= Birthday
			.Parameters("@Gender").Value 		= Gender
			.Parameters("@Mobile").Value 		= Mobile
			.Parameters("@HopePart").Value 		= HopePart
			.Parameters("@HopePart2").Value 		= HopePart2
			.Parameters("@AddrSido").Value 		= AddrSido
			.Parameters("@AddrGugun").Value 	= AddrGugun
			.Parameters("@Addr").Value 			= Addr
			.Parameters("@AddrDetail").Value 		= AddrDetail
			.Parameters("@Email").Value 			= Email
			.Parameters("@Zipcode1").Value 		= Zipcode1
			.Parameters("@Zipcode2").Value 		= Zipcode2
			.Parameters("@Jumin1").Value 		= Jumin1
			.Parameters("@Jumin2").Value 		= Jumin2
			.Parameters("@EzCareer").Value 		= EzCareer
			
			.Parameters("@Intro").Value 			= Intro

			.Parameters("@EduLast").Value 			= EduLast
			.Parameters("@EduType").Value 			= EduType
			
			.Parameters("@ImgPath").Value 		= ImgPath

			.Parameters("@Attach1_Path").Value 	= Attach1_Path
			.Parameters("@Attach1_FName").Value = Attach1_Name
			.Parameters("@Attach2_Path").Value 	= Attach2_Path
			.Parameters("@Attach2_FName").Value = Attach2_Name
			.Parameters("@Attach3_Path").Value 	= Attach3_Path
			.Parameters("@Attach3_FName").Value = Attach3_Name

			.Parameters("@DelImgFlag").Value 	= Fn_SetDefault(DelImgPath,"","0","1")
			.Parameters("@DelAtt1Flag").Value 	= Fn_SetDefault(DelAttach_1,"","0","1")
			.Parameters("@DelAtt2Flag").Value 	= Fn_SetDefault(DelAttach_2,"","0","1")
			.Parameters("@DelAtt3Flag").Value 	= Fn_SetDefault(DelAttach_3,"","0","1")

			Set Rs = .Execute

			'Result = .Parameters("@Result")
			
			If Not Rs.Eof Then
				Result = Rs("Result")
			Else
				Result = 99
			End If

			If Flag = "ADD" And Result = 0 Then
			
				Set Rs = Rs.NextRecordSet()
				
				If Not Rs.Eof Then
					Seq = Rs("UsrSeq")
				End If
				
			End If
			
		End With
		
		Rs.Close
		Set Rs = Nothing

		
		If Result = 0 And RecrSeq <> "0" And Seq <> "0" Then		Call setRecrProc
		
	End Sub

	Sub setRecrProc()
		Dim tmpResult : tmpResult = 0
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "ADD_ADM"
			.Parameters("@RecrSeq").Value 		= RecrSeq
			.Parameters("@UsrSeq").Value 		= Seq
			
			.Parameters("@PartnerDepth1").Value = PartnerDepth1
			.Parameters("@PartnerDepth2").Value = PartnerDepth2
			
			.Execute
			
			tmpResult = .Parameters("@Result")
			
		End With
		
		If tmpResult <> 0 Then
			Result = 90
		Else
			Result = 0
		End If
	End Sub

	Sub setEduProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUserEdu_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= FlagSub
			.Parameters("@EduSort").Value 	= Sort
			
			.Parameters("@UsrSeq").Value 	= Seq
			.Parameters("@SchoolName").Value = SchoolName
			.Parameters("@EduBYear").Value 	= EduBYear
			.Parameters("@EduBMonth").Value = EduBMonth
			.Parameters("@EduEYear").Value 	= EduEYear
			.Parameters("@EduEMonth").Value = EduEMonth
			.Parameters("@Status").Value 	= Status
			.Parameters("@Local").Value 	= Local
			.Parameters("@Major").Value 	= Major
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub

	Sub setCareerProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUserCareer_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= FlagSub
			.Parameters("@CrSort").Value 	= Sort
			
			.Parameters("@UsrSeq").Value 	= Seq
			.Parameters("@CrBYear").Value 	= CrBYear
			.Parameters("@CrBMonth").Value 	= CrBMonth
			.Parameters("@CrEYear").Value 	= CrEYear
			.Parameters("@CrEMonth").Value 	= CrEMonth
			.Parameters("@Reason").Value 	= Reason
			.Parameters("@CompanyName").Value = CompanyName
			.Parameters("@Department").Value = Department
			.Parameters("@Task").Value 		= Task
			.Parameters("@Position").Value 	= Position
			.Parameters("@CompanyYear").Value 	= CompanyYear
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub	
	
	Sub setFamProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUserFam_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= FlagSub
			.Parameters("@FamSort").Value 	= Sort
			
			.Parameters("@UsrSeq").Value 	= Seq
			.Parameters("@Relation").Value 	= Relation
			.Parameters("@FmName").Value 	= FmName
			.Parameters("@Age").Value 		= Age
			.Parameters("@Education").Value = Education
			.Parameters("@Job").Value 		= Job
			.Parameters("@IsLiveWith").Value = IsLiveWith
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub	

	Sub setMemoProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUserMemo_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= FlagSub
			.Parameters("@UsrMSeq").Value 	= Sort
			
			.Parameters("@UsrSeq").Value 	= Seq
			.Parameters("@AdmSeq").Value 	= Session("ASeq")
			.Parameters("@Memo").Value 		= Memo
			
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
				If Not FN_isBlank(DelImgPath) Then  	Call clsObjUpload.Delete(strPath & "/" & DelImgPath)
				If Not FN_isBlank(DelAttach_1) Then  	Call clsObjUpload.Delete(strPath & "/" & DelAttach_1)
				If Not FN_isBlank(DelAttach_2) Then  	Call clsObjUpload.Delete(strPath & "/" & DelAttach_2)
				If Not FN_isBlank(DelAttach_3) Then  	Call clsObjUpload.Delete(strPath & "/" & DelAttach_3)
				
				strMsg = "처리가 완료됐습니다."
				
				If Flag = "ADD" Then
					strScript = "window.location.href = ""resume.asp""; "
				ElseIf Flag = "MOD" Then
					strScript = "window.location.href = ""resume.asp""; "
				ElseIf Flag = "MEMO" Then
					strScript = "window.location.href = ""mem_memo.asp?Seq=" & Seq & """; "
				Else
					strScript = "window.location.href = ""mem_list.asp?sFlag=" & sFlag & "&sView=" & sView & """; "
				End If
			ElseIf Result = 1 Then
				strMsg = "이미 등록되어 있는 핸드폰번호 입니다."
				strScript = "history.back()"
			ElseIf Result = 90 Then
				strMsg = "회원정보는 정상 처리 되었으나 채용지원 등록중 오류가 발생했습니다."
				strScript = "window.location.href = ""mem_list.asp""; "
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
			<title>회원 정보 처리 - HC 관리자</title>
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
