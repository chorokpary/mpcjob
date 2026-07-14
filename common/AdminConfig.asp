<%
'____________________________________________________________________________________
'
' * Discription : AdminConfig.asp / 관리자 페이지정보 체크
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	'페이지 정보 체크 ++++++++++++++++
	Dim clsMap		: Set clsMap = New ClsFileMap
		
	'페이지 정보 얻기
	clsMap.setInitDicAdmin()
	
	'기본키 : "" / 임의키 : ClsFileMaps.getDicKey
	Dim GNB_STR_MNCD	: GNB_STR_MNCD = clsMap.getMenuCode("")
	Dim GNB_STR_LOC		: GNB_STR_LOC = clsMap.getLocationAdmin("")
	Dim GNB_STR_AUTH	: GNB_STR_AUTH = clsMap.getMenuAuthAdmin("")
	
	Dim GNB_STR_MND1 : GNB_STR_MND1 = Left(GNB_STR_MNCD,2)
	Dim GNB_STR_MND2 : GNB_STR_MND2 = Mid(GNB_STR_MNCD,3,2)
	
	Set clsMap = Nothing
	
	'권한 체크 ++++++++++++++++
	Dim clsAuth		: Set clsAuth = New ClsChkAuth
	
	'로그인 체크
	clsAuth.chkAdmin(GNB_STR_AUTH)


	'GNB 관련 함수 ++++++++++++++++
	' 메뉴 활성화
	Function FN_AdmGnbOn(argMn,argSelMn)
		FN_AdmGnbOn = Fn_SetDefault(argMn,argSelMn,"class=""on""","")
	End Function

	' 메뉴 권한
	Function FN_AdmMenuAuth(argMnAuth)
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
			FN_AdmMenuAuth = True
		Else
			FN_AdmMenuAuth = False
		End If
		
	End Function
	
	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 

	
%>