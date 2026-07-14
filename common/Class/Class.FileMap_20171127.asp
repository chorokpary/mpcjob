<%
'____________________________________________________________________________________
'
' * Discription : Class.FileMap.asp / 파일 정보 Mapping
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-03-21 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


Class ClsFileMap

	Private IsErrDisplay
	Private DicMenuCode

	'========================================================
	' 클래스 생성자.
	'========================================================
    Sub Class_Initialize()
		IsErrDisplay = True
    End Sub


	'========================================================
	' 클래스 소멸자
	'========================================================
    Sub Class_Terminate
    	Set DicMenuCode = Nothing
    End Sub
    

	'========================================================
	' SET Dictionary info initialize
	'========================================================	
	Public Sub setInitDic ()
	
		Set DicMenuCode = CreateObject("Scripting.Dictionary")

		'##### Layer Popup 및 ajax 파일은 별도로 등록하지 않음.(등록되지 않은 파일은 무제한 접근 가능)

		'### 채용정보  ####
		Call DicMenuCode.Add("/recruit/recruit_list.asp", Array("010100","\t","0"))
		Call DicMenuCode.Add("/recruit/recruit_view.asp", Array("010100","\t","0"))
		Call DicMenuCode.Add("/recruit/recruitUp.asp", Array("010100","\t","0"))

		'### HC Familly  ####
		Call DicMenuCode.Add("/family/family.asp", Array("020100","\t","0"))

		'### About HC   ####
		Call DicMenuCode.Add("/corp/philosophy.asp", Array("030100","회사소개\t /corp/philosophy.asp \n 경영철학 \t","0"))
		Call DicMenuCode.Add("/corp/people.asp", Array("030100","회사소개 \t /corp/philosophy.asp \n 인재상 \t","0"))
		Call DicMenuCode.Add("/corp/benefit.asp", Array("030100","회사소개 \t /corp/philosophy.asp \n 급여 및 복리후생 \t","0"))
		Call DicMenuCode.Add("/corp/eduSystem.asp", Array("030100","회사소개 \t /corp/philosophy.asp \n 교육제도 \t","0"))
		Call DicMenuCode.Add("/corp/center.asp", Array("030200","센터찾기 \t","0"))
		
		'### HC 알리미   ####
		Call DicMenuCode.Add("/inform/notice_list.asp", Array("040100","공지사항\t","0"))
		Call DicMenuCode.Add("/inform/notice_view.asp", Array("040100","공지사항\t","0"))
		Call DicMenuCode.Add("/inform/faq_list.asp", Array("040200","FAQ\t","0"))

		'### MyPage   ####
		Call DicMenuCode.Add("/mypage/article.asp", Array("050000","이력서 관리\t","1"))
		Call DicMenuCode.Add("/mypage/article_proc.asp", Array("050000","이력서 관리\t","1"))
		Call DicMenuCode.Add("/mypage/resume.asp", Array("050100","이력서 관리\t","2"))
		Call DicMenuCode.Add("/mypage/resume_modify.asp", Array("050100","이력서 관리\t","2"))
		Call DicMenuCode.Add("/mypage/resume_proc.asp", Array("050100","이력서 관리\t","2"))
		Call DicMenuCode.Add("/mypage/apply.asp", Array("050200","입사지원관리\t","2"))
		Call DicMenuCode.Add("/mypage/apply_delete.asp", Array("050200","입사지원관리\t","2"))
		Call DicMenuCode.Add("/mypage/matchArea.asp", Array("050300","인근지역공고\t","2"))
		Call DicMenuCode.Add("/mypage/memSecession.asp", Array("050400","회원탈퇴\t","1"))
		Call DicMenuCode.Add("/mypage/memSecession_proc.asp", Array("050400","회원탈퇴\t","1"))

		'### 회원관련   ####
		Call DicMenuCode.Add("/membership/login.asp", Array("060100","로그인\t","0"))
		Call DicMenuCode.Add("/membership/find.asp", Array("060100","비밀번호찾기\t","0"))
		
		'### 기타   ####
		Call DicMenuCode.Add("/etc/privacy.asp", Array("070100","개인정보취급방침\t","0"))
		Call DicMenuCode.Add("/etc/service.asp", Array("070200","이용약관\t","0"))
		

	End Sub
	
	
	'========================================================
	' GET Location 
	'========================================================	
	Public Function getLocation(argDicKey)
		Dim strLoc : strLoc = "<a href='/'>HOME</a> : "
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		Dim strMenuCode : strMenuCode = getMenuCode(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			Select Case Left(strMenuCode,2)
				Case "01" :	strLoc = strLoc & "<a href="""">채용정보</a>"
				Case "02" :	strLoc = strLoc & "<a href=""/family/family.asp"">HC Family</a>"
				Case "03" :	strLoc = strLoc & "<a href=""/corp/philosophy.asp"">About HC</a>"
				Case "04" :	strLoc = strLoc & "<a href=""/inform/notice_list.asp"">HC알리미</a>"
				Case "05" :	strLoc = strLoc & "<a href=""/mypage/resume.asp"">MY PAGE</a>"
				Case "06" : strLoc = strLoc & "<a href=""""></a>"
			End Select
			
			strLoc = strLoc & getLocationTag(DicMenuCode(strDicKey)(1))
		End If
		
		getLocation = strLoc
	End Function

	'========================================================
	' GET Menu Auth 
	'========================================================	
	Public Function getMenuAuth(argDicKey)
		Dim strMenuAuth : strMenuAuth = "0" '기본 메뉴권한 : 비로그인
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			strMenuAuth = DicMenuCode(strDicKey)(2)
		End If
		
		getMenuAuth = strMenuAuth
	End Function
	

	'========================================================
	' SET Dictionary info initialize : Admin
	'========================================================	
	Public Sub setInitDicAdmin ()
		Set DicMenuCode = CreateObject("Scripting.Dictionary")
		
		'Call DicMenuCode.Add("url", Array("050100","[위치]\t[URL]\n[위치]\t","1"))
		
		'### 채용관리  ####
		Call DicMenuCode.Add("/admin/recruit/job_skin_list.asp", Array("010100","진행중인 채용공고\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_skin_elist.asp", Array("010200","마감된 채용공고\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_skin_alist.asp", Array("010400","전체 채용공고\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_skin_reg.asp", Array("010300","채용공고 등록\t","2"))
		Call DicMenuCode.Add("/admin/recruit/job_skin_mod.asp", Array("010000","채용공고 수정\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_apply_stats.asp", Array("010000","지원자현황\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_apply_user.asp", Array("010000","지원자관리\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_apply_user_xls.asp", Array("010000","지원자관리\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_apply_proc.asp", Array("010000","지원자관리\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_adm_ajax.asp", Array("010000","채용공고 심사권한\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_partner_stats.asp", Array("010000","통계관리\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_partner_stats_xls.asp", Array("010000","통계관리\t","3"))
		Call DicMenuCode.Add("/admin/recruit/job_recom_list.asp", Array("010000","추천공고 관리\t","1"))
		Call DicMenuCode.Add("/admin/recruit/job_proc.asp", Array("010000","채용공고 관리\t","3"))
		

		'### 회원관리  ####
		Call DicMenuCode.Add("/admin/member/mem_list.asp", Array("020100","회원관리\t","2"))
		Call DicMenuCode.Add("/admin/member/mem_form.asp", Array("020100","회원정보\t","2"))
		Call DicMenuCode.Add("/admin/member/mem_info_ajax.asp", Array("020100","회원정보\t","2"))
		Call DicMenuCode.Add("/admin/member/mem_xls.asp", Array("020100","회원관리\t","2"))
		Call DicMenuCode.Add("/admin/member/mem_print.asp", Array("020100","회원관리\t","3"))
		Call DicMenuCode.Add("/admin/member/mem_proc.asp", Array("020100","회원관리\t","2"))
		Call DicMenuCode.Add("/admin/member/mem_memo.asp", Array("020000","회원관리\t","3"))
		Call DicMenuCode.Add("/admin/member/mem_apply.asp", Array("020000","회원관리\t","3"))

	
		'### 프로젝트관리  ####
		Call DicMenuCode.Add("/admin/project/pjt_list.asp", Array("030100","","1"))
		Call DicMenuCode.Add("/admin/project/pjt_form.asp", Array("030100","","1"))
		Call DicMenuCode.Add("/admin/project/pjt_proc.asp", Array("030100","","1"))
		

		'### 게시글관리  ####
		Call DicMenuCode.Add("/admin/board/notice_list.asp", Array("040100","공지사항\t","1"))
		Call DicMenuCode.Add("/admin/board/notice_form.asp", Array("040100","공지사항\t","1"))
		Call DicMenuCode.Add("/admin/board/notice_proc.asp", Array("040100","공지사항\t","1"))

		Call DicMenuCode.Add("/admin/board/faq_list.asp", Array("040200","FAQ\t","1"))
		Call DicMenuCode.Add("/admin/board/faq_form.asp", Array("040200","FAQ\t","1"))
		Call DicMenuCode.Add("/admin/board/faq_proc.asp", Array("040200","FAQ\t","1"))

		Call DicMenuCode.Add("/admin/board/mail_list.asp", Array("040300","메일링 서비스\t","1"))
		Call DicMenuCode.Add("/admin/board/mail_form.asp", Array("040300","SMS 발송\t","1"))

		Call DicMenuCode.Add("/admin/board/pop_list.asp", Array("040400","팝업관리\t","1"))
		Call DicMenuCode.Add("/admin/board/pop_form.asp", Array("040400","팝업관리\t","1"))
		Call DicMenuCode.Add("/admin/board/pop_proc.asp", Array("040400","팝업관리\t","1"))


		'### HC통계  ####
		Call DicMenuCode.Add("/admin/stats/stats.asp", Array("050100","HC회원 통계\t","1"))
		Call DicMenuCode.Add("/admin/stats/stats_xls.asp", Array("050100","HC회원 통계\t","1"))
		Call DicMenuCode.Add("/admin/stats/stats_detail.asp", Array("050100","HC회원 통계\t","1"))
		Call DicMenuCode.Add("/admin/stats/stats_detail_xls.asp", Array("050100","HC회원 통계\t","1"))

	
		'### 환경설정  ####
		Call DicMenuCode.Add("/admin/siteconf/adm_list.asp", Array("060100","관리자등록\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/adm_form.asp", Array("060100","관리자등록\t","3"))
		Call DicMenuCode.Add("/admin/siteconf/adm_proc.asp", Array("060100","관리자등록\t","3"))
		Call DicMenuCode.Add("/admin/siteconf/conf_main.asp", Array("060200","메인설정\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/conf_main_proc.asp", Array("060200","메인설정\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/conf_flash.asp", Array("060200","메인 플래시 설정\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/conf_flash_form.asp", Array("060200","메인 플래시 설정\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/conf_flash_proc.asp", Array("060200","메인 플래시 설정\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/partner_skin_clist.asp", Array("060300","인력업체 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/partner_skin_olist.asp", Array("060400","온라인 광고 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/partner_form.asp", Array("060000","\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/partner_proc.asp", Array("060000","\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/hiretype_list.asp", Array("060500","고용형태 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/hiretype_form.asp", Array("060500","고용형태 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/hiretype_proc.asp", Array("060500","고용형태 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/recrpart_list.asp", Array("060600","직종 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/recrpart_form.asp", Array("060600","직종 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/recrpart_proc.asp", Array("060600","직종 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/ezcareer_list.asp", Array("060700","경력 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/ezcareer_form.asp", Array("060700","경력 관리\t","1"))
		Call DicMenuCode.Add("/admin/siteconf/ezcareer_proc.asp", Array("060700","경력 관리\t","1"))

		'### 쪽지함  ####
		Call DicMenuCode.Add("/admin/message/message_list.asp", Array("070100","","1"))
		Call DicMenuCode.Add("/admin/message/message_view.asp", Array("070100","","1"))
		Call DicMenuCode.Add("/admin/message/message_proc.asp", Array("070100","","1"))

	End Sub


	'========================================================
	' GET Location : Admin
	'========================================================	
	Public Function getLocationAdmin(argDicKey)
		Dim strLoc : strLoc = "<a href=""/admin/"" class=""home"">HOME</a>"
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		Dim strMenuCode : strMenuCode = getMenuCode(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			Select Case Left(strMenuCode,2)
				Case "01" :	strLoc = strLoc & "<a href=""/admin/recruit/job_list.asp"">채용관리</a>"
				Case "02" :	strLoc = strLoc & "<a href=""/admin/member/mem_list.asp"">회원관리</a>"
				Case "03" :	strLoc = strLoc & "<a href=""/admin/project/pjt_list.asp"">프로젝트관리</a>"
				Case "04" :	strLoc = strLoc & "<a href=""/admin/board/notice_list.asp"">게시글관리</a>"
				Case "05" :	strLoc = strLoc & "<a href=""/admin/stats/stats.asp"">HC통계</a>"
				Case "06" :	strLoc = strLoc & "<a href=""/admin/siteconf/adm_list.asp"">환경설정</a>"
				Case "07" :	strLoc = strLoc & "<a href=""/customer/faq.asp"">쪽지함</a>"
			End Select
			
			strLoc = strLoc & getLocationAdminTag(DicMenuCode(strDicKey)(1))
		End If
		
		getLocationAdmin = strLoc
	End Function

	'========================================================
	' GET Menu Auth : Admin
	'========================================================	
	Public Function getMenuAuthAdmin(argDicKey)
		Dim strMenuAuth : strMenuAuth = "3"	' 메뉴 기본 접근 권한 : 일반관리자 이상 접근가능
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			strMenuAuth = DicMenuCode(strDicKey)(2)
		End If
		
		getMenuAuthAdmin = strMenuAuth
	End Function
	

	'========================================================
	' GET Location Text To Html Tag : Front
	'========================================================	
	Public Function getLocationTag(argLoc)
		Dim strRtn : strRtn = ""
		If Not FN_isBlank(argLoc) Then
			Dim arrDep : arrDep = Split(argLoc,"\n")
			Dim i, arrLink
			
			For i = 0 To UBound(arrDep)
				arrLink = Split(arrDep(i),"\t")
				If IsArray(arrLink) Then
					If FN_isBlank(arrLink(1)) Then
						If arrLink(0) <> "" Then
							strRtn = strRtn & " > " & arrLink(0)
						End If
					Else
						strRtn = strRtn & " > <a href=""" & arrLink(1) & """>" & arrLink(0) & "</a>"
					End If
				Else
					strRtn = strRtn & " > " & arrDep(i)
				End If
			Next
		End If
		
		getLocationTag = strRtn
	End Function

	'========================================================
	' GET Location Text To Html Tag : Admin
	'========================================================	
	Public Function getLocationAdminTag(argLoc)
		Dim strRtn : strRtn = ""
		If Not FN_isBlank(argLoc) Then
			Dim arrDep : arrDep = Split(argLoc,"\n")
			Dim i, arrLink
			
			For i = 0 To UBound(arrDep)
				arrLink = Split(arrDep(i),"\t")
				If IsArray(arrLink) Then
					If FN_isBlank(arrLink(1)) Then
						If arrLink(0) <> "" Then
							strRtn = strRtn & "<span>" & arrLink(0) & "</span>"
						End If
					Else
						strRtn = strRtn & "<a href=""" & arrLink(1) & """>" & arrLink(0) & "</a>"
					End If
				Else
					strRtn = strRtn & "<span>" & arrDep(i) & "</span>"
				End If
			Next
		End If
		
		getLocationAdminTag = strRtn
	End Function
	

	'========================================================
	' GET Dictionary Key 
	'========================================================	
	Public Function getDicKey()
		Dim strDicKey
		Dim strCurPage 		: strCurPage = Request.ServerVariables("PATH_INFO")
		Dim strCurVFlag 	: strCurVFlag = Request("VFlag")
		
		Select Case strCurPage
		    Case "/tech/paper_list.asp", "/investment/statement.asp" : If strCurVFlag = "" Then 	strCurVFlag = "1"
		End Select 
		
		If strCurVFlag <> "" Then
			strDicKey = strCurPage & "#" & strCurVFlag
		Else
			strDicKey = strCurPage
		End If
		
		getDicKey = strDicKey
	End Function


	'========================================================
	' GET Menu Code 
	'========================================================	
	Public Function getMenuCode(argDicKey)
		Dim strMenuCode : strMenuCode = "000000"
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			strMenuCode = DicMenuCode(strDicKey)(0)
		End If
		
		getMenuCode = strMenuCode
	End Function

	'========================================================
	' GET Menu Name 
	'========================================================	
	Public Function getMenuName(argDicKey)
		Dim strMenuName : strMenuName = "메인"
		Dim strDicKey	: strDicKey = chkDicKey(argDicKey)
		
		If IsArray(DicMenuCode(strDicKey)) Then
			strMenuName = DicMenuCode(strDicKey)(2)
		End If
		
		getMenuName = strMenuName
	End Function


	'========================================================
	' Check Dictionary Key : 비어있을 경우 현재 페이지 URL 반환
	'========================================================	
	Public Function chkDicKey(argDicKey)
		If argDicKey = "" Then
			chkDicKey = getDicKey()
		Else
			chkDicKey = argDicKey
		End If
	End Function



	'========================================================
	' 에러정보 출력
	'========================================================	
	Public Sub SetErrDisplay(ByVal as_Boolean)
		SetErrDisplay = Cbool(as_Boolean)
	End Sub
    
End Class
%>