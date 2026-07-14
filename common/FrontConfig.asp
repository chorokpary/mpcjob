<%
'____________________________________________________________________________________
'
' * Discription : FrontConfig.asp / Front 페이지정보 체크
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	'페이지 정보 체크 ++++++++++++++++
	Dim clsMap		: Set clsMap = New ClsFileMap
		
	'페이지 정보 얻기
	clsMap.setInitDic()
	
	'기본키 : "" / 임의키 : ClsFileMaps.getDicKey
	Dim GNB_STR_MNCD	: GNB_STR_MNCD = clsMap.getMenuCode("")
	Dim GNB_STR_LOC		: GNB_STR_LOC = clsMap.getLocation("")
	Dim GNB_STR_AUTH	: GNB_STR_AUTH = clsMap.getMenuAuth("")
	
	Dim GNB_STR_MND1 : GNB_STR_MND1 = Left(GNB_STR_MNCD,2)
	Dim GNB_STR_MND2 : GNB_STR_MND2 = Mid(GNB_STR_MNCD,3,2)
	
	Dim GNB_STR_GREETING : GNB_STR_GREETING = ""

	Set clsMap = Nothing

	'권한 체크 ++++++++++++++++
	Dim clsAuth		: Set clsAuth = New ClsChkAuth
		
	'로그인 체크
	clsAuth.chkLogin(GNB_STR_AUTH)
	
	Set clsAuth = Nothing
	
	'Site Conf 로딩
	Call SB_FrontConfig


	' -- 채용정보 검색용
	Dim GLOBAL_SEARCH_PART : GLOBAL_SEARCH_PART = FN_Req("sPart",FN_Req("Search_Part",""))
	Dim GLOBAL_SEARCH_SIDO : GLOBAL_SEARCH_SIDO = FN_Req("sAddrSido",FN_Req("Search_AddrSido",""))
	Dim GLOBAL_SEARCH_GUGUN : GLOBAL_SEARCH_GUGUN = FN_Req("sAddrGugun",FN_Req("Search_AddrGugun",""))

	'GNB 관련 함수 ++++++++++++++++
	' 메뉴 활성화
	Function FN_FrontGnbOn(argMn,argSelMn)
		FN_FrontGnbOn = Fn_SetDefault(argMn,argSelMn,"class=""on""","")
	End Function

	' 메뉴 활성화
	Function FN_FrontImgOn(argMn,argSelMn)
		FN_FrontImgOn = Fn_SetDefault(argMn,argSelMn,"on","off")
	End Function

	' 메뉴 권한
	Function FN_FrontMenuAuth(argMnAuth)
		Dim AdmLevel, MnLevel
		
		If IsNumeric(Session("ALevel")) Then	
			AdmLevel = CInt(Session("ALevel"))
		Else
			AdmLevel = 0
		End If

		If IsNumeric(argMnAuth) Then	
			MnLevel = CInt(argMnAuth)
		Else
			MnLevel = 0
		End If
		
		If AdmLevel > 0 And MnLevel > 0 And  AdmLevel <= MnLevel Then
			FN_FrontMenuAuth = True
		Else
			FN_FrontMenuAuth = False
		End If
		
	End Function

	' 사이트 설정 정보
	Sub SB_FrontConfig()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfSite_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "INFO"
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				GNB_STR_GREETING = Rs("ConfGreeting")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub

	' 로그인시 접근 제한 페이지
	Sub SB_NoLoginPage()
		'이미 로그인 한 상태일 경우 메인페이지로 이동
		If Session("USeq") <> "" Then
			%><meta http-equiv="refresh" content="0; url=/"><%
			Response.ENd
		End If
	End Sub	

	' 로그인 체크 페이지
	Sub SB_LoginChkPage()
		'이미 로그인 한 상태일 경우 메인페이지로 이동
		If Session("USeq") = "" Then
			%>
			<script type="text/javascript">
				alert('잘못된 접근 또는 로그인 세션이 만료 되었습니다.\n로그인 후 다시 접근하시기 바랍니다.');
				window.location.href = "/membership/login.asp";
			</script>
			<%
			Response.ENd
		End If
	End Sub	


	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 

%>