<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_proc.asp / 회원 정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4m / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, FlagSub, Seq, sFlag, sView, IsAjax, Filter
	Dim PartnerDepth1, PartnerDepth2, RecrSeq, RecrStatus
	Dim UsrName, Birthday, Gender, Mobile, ChkMobile, HopePart, HopePart2, AddrSido, AddrGugun, Addr, AddrDetail
	Dim Jumin1, Jumin2, EzCareer
	Dim Email, Zipcode1, Zipcode2, Intro, EduLast, EduType
	Dim Passwd
	Dim ImgPath, ImgName
	Dim Attach1_Path, Attach2_Path, Attach3_Path
	Dim Attach1_Name, Attach2_Name, Attach3_Name
	Dim Birthday_Y
	
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

		
		If Flag = "DEL" Or Flag = "BLK" Or Flag = "SEARCH" Or Flag = "RECR" Or Flag = "PWD_C" Or Flag = "PWD_S" _
			Or Flag = "EDU" Or Flag = "CAREER" Or Flag = "FAM" Or Flag = "MEMO" Or Flag = "QUES" Or Flag = "AGREE"  Then
			' DEL : 관리자삭제 / BLK : 블랙리스트 처리 / SEARCH : 서칭거부리스트 처리 / RECR : 채용공고 지원 / PWD_C : 비밀번호변경 / PWD_S : 임시비밀번호 전송
			' EDU : 학력정보 / CAREER : 경력정보 / FAM : 가족사항 / MEMO : 관리자 메모 / QUES : 사전질문 답변
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
		ElseIf Flag = "QUES" Then
			Call setQuestionProc()
		Else
			Call setProc()
		End If
		
		Call setResult()
	End Sub
	
	Sub getParamReq()

		Flag 		= FN_Req("Flag","")
		FlagSub		= FN_Req("FlagSub","")
		Seq 		= FN_Req("Seq","0")
		IsAjax 		= FN_Req("IsAjax","0")

		sFlag 		= FN_Req("sFlag","")
		sView 		= FN_Req("sView","0")
		
		Mobile 		= FN_Req("Mobile","0")

		PartnerDepth1 	= FN_Req("PartnerDepth1","")
		PartnerDepth2 	= FN_Req("PartnerDepth2","")
		RecrSeq 		= FN_Req("RecrSeq","0")
		RecrStatus 	= FN_Req("RecrStatus","00")

		Passwd 		= FN_Req("Passwd","")
		
		If Passwd <> "" Then		Passwd = hex_sha1(Passwd)
		
		If Flag = "PWD_C" Or Flag = "PWD_S" Then
			strParamFlag = "@Seq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		Else	' 복수 변경이 가능한 Flag들..
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		End If

		
		Sort = FN_Req("Sort","")
		Sort = Replace(Sort," ","")
		
	End Sub

	Sub getParamObj()

		With clsObjUpload
			Flag 		= FN_StrInj(.getReq("Flag"))
			Seq 		= FN_StrInj(.getReq("Seq"))
			IsAjax 		= FN_StrInj(.getReq("IsAjax"))

			Filter = FN_StrInj(.getReq("Filter"))
		
				
			'If Flag = "ADD" Then
			'	PartnerDepth1 	= FN_StrInj(.getReq("AddPartnerD1"))
			'	PartnerDepth2 	= FN_StrInj(.getReq("AddPartnerD2"))
			'	RecrSeq 		= FN_StrInj(.getReq("AddRecrSeq"))
			'ElseIf Flag = "MOD" Then
			'	PartnerDepth1 	= FN_StrInj(.getReq("PartnerDepth1"))
			'	PartnerDepth2 	= FN_StrInj(.getReq("PartnerDepth2"))
			'	RecrSeq 		= FN_StrInj(.getReq("RecrSeq"))
			'End If 
			PartnerDepth1 	= FN_StrInj(.getReq("PartnerDepth1"))
			PartnerDepth2 	= FN_StrInj(.getReq("PartnerDepth2"))
			RecrSeq 		= FN_StrInj(.getReq("RecrSeq"))
			RecrStatus 	= FN_StrInj(.getReq("RecrStatus"))

			If RecrSeq = "" Then 	RecrSeq = "0"
			If RecrStatus = "" Then 	RecrStatus = "00"
	
			UsrName 	= FN_StrInj(.getReq("UsrName"))
			Birthday 	= FN_StrInj(.getReq("Birthday"))
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
			
			Intro 	= FN_StrInj(.getReq("Intro"))

			EduLast  	= FN_StrInj(.getReq("EduLast"))
			EduType  	= FN_StrInj(.getReq("EduType"))
			
			DelImgPath = FN_StrInj(.getReq("DelImgPath"))
			DelAttach_1 = FN_StrInj(.getReq("DelAttach_1"))
			DelAttach_2 = FN_StrInj(.getReq("DelAttach_2"))
			DelAttach_3 = FN_StrInj(.getReq("DelAttach_3"))

			Birthday_Y = FN_StrInj(.getReq("Birthday_Y"))
			
			If Flag = "ADD" And Addr = "" Then
				Addr = Trim(AddrSido & " " & AddrDetail)
			End If
			
			'// ### Img File Upload : Start
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
			'// ### Img File Upload : End
			
			'// ### Attach File Upload : Start
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
			'// ### Attach File Upload : End
			
		End With
	

		strParamFlag = "@Seq"
		If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
		If Mobile = "" Or ChkMobile <> "Y" Then 	Call RtnErrorObj("[PARAM] 잘못된 경로로 접근하셨습니다.")
		
		' 사용자가 년도를 바르게 입력했을경우 대비 Left 2자리는 보존
		If Birthday_Y > Year(Date()) Or Not IsNumeric(Birthday_Y) Then
			'Birthday_Y = "19"
			Birthday_Y = Left(Birthday_Y,2)
		Else
			Birthday_Y = Left(Birthday_Y,2)
		End If
		
		If Len(Birthday) = 6 Then
			Birthday = Birthday_Y & LEFT(Birthday,2) & "-" & MID(Birthday,3,2) & "-" & MID(Birthday,5,2)
			If Not IsDate(Birthday) Then 		Birthday = ""
		End If

		
	End Sub
		
	Sub setProc()
		Dim tmpPasswd
		
		If Flag = "PWD_S" Then
			tmpPasswd = Fn_EncryptData("NEWPWD")
			tmpPasswd = LEFT(tmpPasswd,6)
			tmpPasswd = UCase(tmpPasswd)
				
			Passwd = hex_sha1(tmpPasswd)
		End If

		Dim Rs


		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters(strParamFlag).Value 	= Seq
			
			.Parameters("@Passwd").Value 	= Passwd
			.Parameters("@UsrName").Value 	= UsrName
			.Parameters("@Birthday").Value 	= Birthday
			.Parameters("@Gender").Value 	= Gender
			.Parameters("@Mobile").Value 	= Mobile
			.Parameters("@HopePart").Value 	= HopePart
			.Parameters("@HopePart2").Value 	= HopePart2
			.Parameters("@AddrSido").Value 	= AddrSido
			.Parameters("@AddrGugun").Value = AddrGugun
			.Parameters("@Addr").Value 		= Addr
			.Parameters("@AddrDetail").Value 	= AddrDetail
			.Parameters("@Email").Value 		= Email
			.Parameters("@Zipcode1").Value 	= Zipcode1
			.Parameters("@Zipcode2").Value 	= Zipcode2
			.Parameters("@Jumin1").Value 	= Jumin1
			.Parameters("@Jumin2").Value 	= Jumin2
			.Parameters("@EzCareer").Value 	= EzCareer
			
			.Parameters("@Intro").Value 		= Intro

			.Parameters("@EduLast").Value 			= EduLast
			.Parameters("@EduType").Value 			= EduType
			
			.Parameters("@ImgPath").Value 	= ImgPath

			.Parameters("@Attach1_Path").Value 	= Attach1_Path
			.Parameters("@Attach1_FName").Value = Attach1_Name
			.Parameters("@Attach2_Path").Value 	= Attach2_Path
			.Parameters("@Attach2_FName").Value = Attach2_Name
			.Parameters("@Attach3_Path").Value 	= Attach3_Path
			.Parameters("@Attach3_FName").Value = Attach3_Name

			.Parameters("@DelImgFlag").Value = Fn_SetDefault(DelImgPath,"","0","1")
			.Parameters("@DelAtt1Flag").Value = Fn_SetDefault(DelAttach_1,"","0","1")
			.Parameters("@DelAtt2Flag").Value = Fn_SetDefault(DelAttach_2,"","0","1")
			.Parameters("@DelAtt3Flag").Value = Fn_SetDefault(DelAttach_3,"","0","1")

			.Parameters("@RecrSeq").Value = RecrSeq

			Set Rs = .Execute

			'Result = .Parameters("@Result")
			
			If Not Rs.Eof Then
				Result = Rs("Result")
				
				' 임시비번 전송 상태에 업데이트 결과가 성공일 경우 SMS 발송
				If Flag = "PWD_S" AND Result = 0 Then
					Call SB_Send_Emfo(Replace(Mobile,"-",""), Seq, "SMS","[관리자-임시비밀번호 전송] " & Mobile & " 회원", "#MPCJOB# 임시 비밀번호를 전송합니다. [" & tmpPasswd & "]","N", "" ,"" ,"", "ADMPWD", 0, "N", "", "")
				End If
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


		' 채용공고 지원
		If Result = 0 And RecrSeq <> "0" And Seq <> "0" Then
			Call setRecrProc
		End If
		
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

			.Parameters("@Status").Value = RecrStatus
			
			.Execute
			
			tmpResult = .Parameters("@Result")
			
		End With
		
		If tmpResult <> 0 Then
			Result = 90
		Else
			Result = 0
		End If
	End Sub

	' ### Education : Start
	Sub setEduProc()
		Dim EduBYear, EduBMonth, EduEYear, EduEMonth, SchoolName, Major, Status, Local

		EduBYear 	= FN_Req("EduBYear","")
		EduBMonth 	= FN_Req("EduBMonth","")
		EduEYear 	= FN_Req("EduEYear","")
		EduEMonth 	= FN_Req("EduEMonth","")
		SchoolName 	= FN_Req("SchoolName","")
		Major 		= FN_Req("Major","")
		Status 		= FN_Req("Status","")
		Local 		= FN_Req("Local","")

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
	' ### Education : End

	' ### Career : Start
	Sub setCareerProc()
		Dim CrBYear, CrBMonth, CrEYear, CrEMonth, Reason, CompanyName, Department, Task, Position, CompanyYear

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
	' ### Career : End
	
	' ### Familly : Start
	Sub setFamProc()
		Dim Relation, FmName, Age, Education, Job, IsLiveWith 

		Relation 	= FN_Req("Relation","")
		FmName 		= FN_Req("FmName","")
		Age 		= FN_Req("Age","")
		Education 	= FN_Req("Education","")
		Job 		= FN_Req("Job","")
		IsLiveWith 	= FN_Req("IsLiveWith","")

		If Age = "" Or Not IsNumeric(Age) Then 					Age = Null
		If IsLiveWith = "" Or Not IsNumeric(IsLiveWith) Then 	IsLiveWith = Null

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
	' ### Familly : End

	' ### Memo : Start
	Sub setMemoProc()
		Dim Memo
		
		Memo 	= FN_Req("Memo","")

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
	' ### Memo : End

	' ### Question : Start
	Sub setQuestionProc()
		Dim DtSize
		Dim fnQCode, fnAnswer, fnI
		DtSize = FN_Req("DtSize","0")
				
		If IsNumeric(DtSize) Then
			For fnI = 1 To DtSize
				fnQCode = FN_Req("QCode_"& fnI ,"0")
				fnAnswer = FN_Req("Answer_"& fnI ,"")
				
				If Not IsNumeric(fnQCode) Then	fnQCode = 0
				If fnQCode > 0 Then
					Call setApplyAnswer(fnQCode, fnAnswer)
				End If
			Next
		End If
	End Sub

	Sub setApplyAnswer(argQCode, argAnswer)
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruitQaA_Proc"
			.Parameters.Refresh
			.Parameters("@RecrSeq").Value 		= RecrSeq
			.Parameters("@QCode").Value 		= argQCode
			.Parameters("@Answer").Value 		= argAnswer
			.Parameters("@UsrSeq").Value 		= Seq
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub
	' ### Question : End
	
	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		Dim strPath : strPath = Server.MapPath("/")
		
		If IsAjax = "1" Then
			If Result = 0 Then
				If Flag = "PWD_S" Then
					Response.Write "SUCC"
				Else
					Response.Write "SUCC"
				End If 
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
					strScript = "window.location.href = ""mem_list.asp""; "
				ElseIf Flag = "MOD" Then
					strScript = "window.location.href = ""mem_form.asp?Seq=" & Seq & "&Filter=" & Filter & "&sFlag=" & sFlag & "&sView=" & sView & """; "
				ElseIf Flag = "MEMO" Then
					strScript = "window.location.href = ""mem_memo.asp?Seq=" & Seq & """; "
				ElseIf Flag = "QUES" Then
					strScript = "window.location.href = ""mem_question.asp?Seq=" & Seq & "&RecrSeq=" & RecrSeq& """; "
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
