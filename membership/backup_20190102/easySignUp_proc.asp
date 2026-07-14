<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : easySignUp_proc.asp / 간편가입 처리 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4m / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2015-03-06 / 이길홍 / 4M / 학력사항, 경력사랑, 첨부파일, 자기소개서 등등 기능 추가
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, PageFlag, Seq
	Dim UsrName, Birthday, Gender, Mobile, HopePart, AddrSido, AddrGugun, Addr, EzCareer, RecrSeq
	Dim Passwd

	Dim Email, Zipcode1, Zipcode2, Intro, EduLast, EduType

	Dim CrBYear, CrBMonth, CrEYear, CrEMonth, Reason, CompanyName, Department, Task, Position, CompanyYear

	Dim CompanyYear_1, CompanyYear_2, CompanyYear_3
	Dim CompanyName_1, CompanyName_2, CompanyName_3
	Dim Task_1, Task_2, Task_3

	Dim ImgPath, ImgName
	Dim Attach1_Path, Attach2_Path, Attach3_Path
	Dim Attach1_Name, Attach2_Name, Attach3_Name
	
	Dim DelImgPath, DelAttach_1, DelAttach_2, DelAttach_3
	
	Dim strParamFlag
	
	Dim FilePath_Att 	: FilePath_Att = "profile"
	Dim FilePath_Photo 	: FilePath_Photo = "photo"
	Dim clsObjUpload : clsObjUpload = null
	

	Dim Result
	
	Call Main()
	
	Sub Main()


		PageFlag = FN_Req("PageFlag","")
		
		If PageFlag = "JOIN" Or PageFlag = "NOFILE" Then
			' JOIN : 가입
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
		
		If PageFlag = "JOIN" Then
		Else 
		'개인 정보 동의 처리
		Call setProc_upGrade()

		'경력 사항 처리_1 
		CompanyYear = CompanyYear_1
		CompanyName = CompanyName_1
		Task		= Task_1

		If CompanyName <> "" Then 
			Call setCareerProc()
			CompanyName = ""
		End If 

		'경력 사항 처리_2
		CompanyYear = CompanyYear_2
		CompanyName = CompanyName_2
		Task		= Task_2

		If CompanyName <> "" Then 
			Call setCareerProc()
			CompanyName = ""
		End If 

		'경력 사항 처리_3 
		CompanyYear = CompanyYear_3
		CompanyName = CompanyName_3
		Task		= Task_3

		If CompanyName <> "" Then 
			Call setCareerProc()
			CompanyName = ""
		End If 

		End if

		Call setResult()

	End Sub

	Sub getParamReq()
		
		Flag 		= FN_Req("Flag","")
		PageFlag	= FN_Req("PageFlag","")
		UsrName 	= FN_Req("UsrName","")
		Birthday 	= FN_Req("Birthday_Y","") & "-" & FN_Req("Birthday_M","") & "-" & FN_Req("Birthday_D","")
		Gender 		= FN_Req("Gender","")
		Mobile 		= FN_Req("Mobile_1","") & "-" & FN_Req("Mobile_2","") & "-" & FN_Req("Mobile_3","")
		HopePart 	= FN_Req("HopePart","")
		AddrSido 	= FN_Req("AddrSido","")
		AddrGugun	= FN_Req("AddrGugun","")
		Passwd 		= FN_Req("Passwd","")
		EzCareer	= FN_Req("EzCareer","")

		Intro  		= FN_Req("Intro","")

		EduLast  	= FN_Req("EduLast","0")
		EduType  	= FN_Req("EduType","1")

		CompanyYear_1	= FN_Req("CompanyYear_1","")
		CompanyYear_2	= FN_Req("CompanyYear_2","")
		CompanyYear_3	= FN_Req("CompanyYear_3","")

		CompanyName_1	= FN_Req("CompanyName_1","")
		CompanyName_2	= FN_Req("CompanyName_2","")
		CompanyName_3	= FN_Req("CompanyName_3","")

		Task_1	= FN_Req("Task_1","")
		Task_2	= FN_Req("Task_2","")
		Task_3	= FN_Req("Task_3","")

		
		RecrSeq 	= FN_Req("RecrSeq","0")
		
		If Passwd <> "" Then		Passwd = hex_sha1(Passwd)

		Addr = AddrSido & " " & AddrGugun

		If Not IsNumeric(RecrSeq) Then RecrSeq = 0

	End Sub

	
	Sub getParamObj()

		With clsObjUpload
		
			Flag 		= FN_StrInj(.getReq("Flag"))
			PageFlag	= FN_StrInj(.getReq("PageFlag"))
			UsrName 	= FN_StrInj(.getReq("UsrName"))
			Birthday 	= FN_StrInj(.getReq("Birthday_Y")) & "-" & FN_StrInj(.getReq("Birthday_M")) & "-" & FN_StrInj(.getReq("Birthday_D"))
			Gender 		= FN_StrInj(.getReq("Gender"))
			Mobile 		= FN_StrInj(.getReq("Mobile_1")) & "-" & FN_StrInj(.getReq("Mobile_2")) & "-" & FN_StrInj(.getReq("Mobile_3"))
			HopePart 	= FN_StrInj(.getReq("HopePart"))
			AddrSido 	= FN_StrInj(.getReq("AddrSido"))
			AddrGugun	= FN_StrInj(.getReq("AddrGugun"))
			Email 		= FN_StrInj(.getReq("Email"))
			Passwd 		= FN_StrInj(.getReq("Passwd"))
			EzCareer	= FN_StrInj(.getReq("EzCareer"))

			Intro  		= FN_StrInj(.getReq("Intro"))

			EduLast  	= FN_StrInj(.getReq("EduLast"))
			EduType  	= FN_StrInj(.getReq("EduType"))
			
			Addr = Trim(AddrSido & " " & AddrGugun)

			CompanyYear_1	= FN_StrInj(.getReq("CompanyYear_1"))
			CompanyYear_2	= FN_StrInj(.getReq("CompanyYear_2"))
			CompanyYear_3	= FN_StrInj(.getReq("CompanyYear_3"))

			CompanyName_1	= FN_StrInj(.getReq("CompanyName_1"))
			CompanyName_2	= FN_StrInj(.getReq("CompanyName_2"))
			CompanyName_3	= FN_StrInj(.getReq("CompanyName_3"))

			Task_1	= FN_StrInj(.getReq("Task_1"))
			Task_2	= FN_StrInj(.getReq("Task_2"))
			Task_3	= FN_StrInj(.getReq("Task_3"))

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

			// ### Attach File Upload : End


			RecrSeq 	= FN_StrInj(.getReq("RecrSeq"))

			If Passwd <> "" Then		Passwd = hex_sha1(Passwd)
			If Not IsNumeric(RecrSeq) Then RecrSeq = 0

			
			
		End With
	

		
	End Sub


	Sub setProc()

		Set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters("@Seq").Value 			= 0
			
			.Parameters("@Passwd").Value 		= Passwd
			.Parameters("@UsrName").Value 		= UsrName
			.Parameters("@Birthday").Value 		= Birthday
			.Parameters("@Gender").Value 		= Gender
			.Parameters("@Mobile").Value 		= oSeed.Encrypt(Mobile)
			.Parameters("@HopePart").Value 		= HopePart
			.Parameters("@AddrSido").Value 		= AddrSido
			.Parameters("@AddrGugun").Value 	= AddrGugun
			.Parameters("@Addr").Value 			= Addr
			.Parameters("@Email").Value 		= Email
			.Parameters("@EzCareer").Value 		= EzCareer

			.Parameters("@Intro").Value 		= Intro

			.Parameters("@EduLast").Value 		= EduLast
			.Parameters("@EduType").Value 		= EduType

			.Parameters("@ImgPath").Value 		= ImgPath

			.Parameters("@Attach1_Path").Value 	= Attach1_Path
			.Parameters("@Attach1_FName").Value = Attach1_Name
			.Parameters("@Attach2_Path").Value 	= Attach2_Path
			.Parameters("@Attach2_FName").Value = Attach2_Name
			.Parameters("@Attach3_Path").Value 	= Attach3_Path
			.Parameters("@Attach3_FName").Value = Attach3_Name

			.Parameters("@RecrSeq").Value 		= RecrSeq
			
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
					Session("USeq") = Rs("UsrSeq")
					Session("UID")	= oSeed.Decrypt(Rs("Mobile"))
					Session("UName") = Rs("UsrName")
					Session("UType") = Rs("UsrType")
					
					Seq = Rs("UsrSeq")
				End If
				
			End If
			
		End With
		
		Rs.Close
		Set Rs = Nothing



	End Sub

	'개인 정보 동의 처리
	Sub setProc_upGrade()
		Dim Rs_up
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "UPGRADE"
			.Parameters("@Seq").Value 		= Session("USeq")

			Set Rs_up = .Execute
			
			If Not Rs_up.Eof Then
				Result = Rs_up("Result")
			Else
				Result = 99
			End If
			
		End With
		
		Set Rs_up = Nothing

		If Result = "0" Then
			Session("UType") = 1
		End If
		
		Set Rs_up = Nothing
		
	End Sub


	Sub setCareerProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUserCareer_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADD"
			.Parameters("@CrSort").Value 	= ""
			
			.Parameters("@UsrSeq").Value 	= Session("USeq")
			.Parameters("@CrBYear").Value 	= ""
			.Parameters("@CrBMonth").Value 	= ""
			.Parameters("@CrEYear").Value 	= ""
			.Parameters("@CrEMonth").Value 	= ""
			.Parameters("@Reason").Value 	= ""
			.Parameters("@CompanyName").Value = CompanyName
			.Parameters("@Department").Value = ""
			.Parameters("@Task").Value 		= Task
			.Parameters("@Position").Value 	= ""
			.Parameters("@CompanyYear").Value 	= CompanyYear
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub	



	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			Response.Write "SUCC"
		ElseIf Result = 1 Then
			Response.Write "EXISTS"
		Else 
			Response.Write "FAIL"
		End If
		Response.End

	End Sub


	Sub RtnErrorObj (argMsg) 
		Set clsObjUpload = Nothing
		Call SB_ReturnErr(argMsg,"BACK")
	End Sub	
%>
