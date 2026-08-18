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

	' 감사 로그 단순 GET 조회 허용 여부 (True: 조회 로그 기록, False: 기록 생략)
	Dim g_AllowGetReadLog : g_AllowGetReadLog = True

	
%>
<!-- #include virtual="/common/Function/FnAuditLog.asp" -->
<%
'====================================================================================
' * Function : AutoDetectAuditLog
' * Description : 전역 요청 파라미터 분석 및 고도화 자동 감지 엔진
'====================================================================================
Sub AutoDetectAuditLog()
    On Error Resume Next

    ' 1. 세션 관리자 정보가 없으면 (로그인 전 페이지 등) 스킵
    If Session("AID") & "" = "" Then Exit Sub

    Dim scriptName, requestMethod, actionParam, logType, targetKey, logDesc
    Dim itemKey, itemVal, formParamStr, queryParamStr, paramKeyLower, actionKorName
    Dim isProcPage, isReadPage, reqMethodUpper
    Dim isMultipart, isExcelReq

    scriptName    = LCase(Request.ServerVariables("SCRIPT_NAME")) ' 요청 URL (예: /admin/siteconf/iplimit/ip_proc.asp)
    requestMethod = Request.ServerVariables("REQUEST_METHOD")     ' GET 또는 POST
    reqMethodUpper = UCase(requestMethod)
    
    ' mem_proc.asp, pjt_proc.asp 등 자체 명시적 감사로그를 기록하는 페이지는 자동 감지 스킵 (중복 적재 방지)
    If InStr(scriptName, "mem_proc.asp") > 0 Or InStr(scriptName, "pjt_proc.asp") > 0 Then Exit Sub
    
    ' Content-Type 확인하여 멀티파트(파일 업로드) 요청 여부 판별
    isMultipart = (InStr(LCase(Request.ServerVariables("CONTENT_TYPE")), "multipart/form-data") > 0)

    ' 2. 엑셀 및 파일 다운로드 패턴 감지 및 행위 라벨링
    If isMultipart Then
        isExcelReq = (Request.QueryString("excel") = "Y")
    Else
        isExcelReq = (Request("excel") = "Y")
    End If

    If InStr(scriptName, "xls") > 0 Or InStr(scriptName, "excel") > 0 Or isExcelReq Then
        logType = "EXCEL"
        actionKorName = "엑셀 다운로드"
    ElseIf InStr(scriptName, "stats") > 0 Then
        logType = "READ"
        actionKorName = "통계 조회/검색"
    ElseIf InStr(scriptName, "down") > 0 Or InStr(scriptName, "download") > 0 Then        
        If InStr(scriptName, "down_list") > 0 Then
            logType = "READ"
            actionKorName = "다운로드 목록 조회"
        Else            
            logType = "FILEDOWN"
            actionKorName = "파일 다운로드"
        End If
    Else
        ' 3. C/U/D 처리 파일(*_proc.asp, *_ajax.asp) 또는 POST 요청 시 행위 감지
        isProcPage = (InStr(scriptName, "_proc") > 0 Or InStr(scriptName, "_ajax") > 0)
        isReadPage = (InStr(scriptName, "list.asp") > 0 _
                     Or InStr(scriptName, "view.asp") > 0 _
                     Or InStr(scriptName, "search.asp") > 0 _
                     Or InStr(scriptName, "pop.asp") > 0 _
                     Or InStr(scriptName, "reg.asp") > 0 _
                     Or InStr(scriptName, "mod.asp") > 0 _
                     Or InStr(scriptName, "form.asp") > 0 _
                     Or InStr(scriptName, "conf_main.asp") > 0 _
                     Or InStr(scriptName, "job_apply_user.asp") > 0)

        If isProcPage Or (reqMethodUpper = "POST" And Not isReadPage) Then
            ' 폼/쿼리 파라미터(action, Flag, mode, act) 분석 (멀티파트의 경우 Request("Flag") 등 호출 시 BinaryRead 에러 방지를 위해 QueryString만 분석)
            If isMultipart Then
                actionParam = UCase(Trim(Request.QueryString("action") & Request.QueryString("Flag") & Request.QueryString("mode") & Request.QueryString("act")))
            Else
                actionParam = UCase(Trim(Request("action") & Request("Flag") & Request("mode") & Request("act")))
            End If

            If InStr(actionParam, "INSERT") > 0 Or InStr(actionParam, "ADD") > 0 Or InStr(actionParam, "REG") > 0 Or InStr(actionParam, "WRITE") > 0 Then
                logType = "CREATE"
                actionKorName = "데이터 등록"
            ElseIf InStr(actionParam, "DELETE") > 0 Or InStr(actionParam, "DEL") > 0 Or InStr(actionParam, "REMOVE") > 0 Then
                logType = "DELETE"
                actionKorName = "데이터 삭제"
            Else
                logType = "UPDATE"
                actionKorName = "데이터 수정/변경"
            End If
        ElseIf isReadPage And (reqMethodUpper = "POST" Or g_AllowGetReadLog) Then
            ' 리스트/뷰 페이지 조회 시 감사 로그 기록 (POST 검색 또는 GET 허용 시)
            logType = "READ"
            If reqMethodUpper = "POST" Then
                actionKorName = "데이터 조회/검색"
            Else
                actionKorName = "데이터 조회/진입"
            End If
        Else
            ' 일반 단순 리스트/폼 조회 페이지는 로그 생성 스킵 (READ 로그 폭주 방지)
            Exit Sub
        End If
    End If

    ' 4. PK 타겟키 자동 추출 (seq, id, idx, pno 파라미터 감지)
    ' 멀티파트 요청 시 Request("seq") 등 호출로 인한 BinaryRead 에러 방지
    If isMultipart Then
        targetKey = Request.QueryString("seq") & ""
        If targetKey = "" Then targetKey = Request.QueryString("Seq") & ""
        If targetKey = "" Then targetKey = Request.QueryString("id") & ""
        If targetKey = "" Then targetKey = Request.QueryString("idx") & ""
    Else
        targetKey = Request("seq") & ""
        If targetKey = "" Then targetKey = Request("Seq") & ""
        If targetKey = "" Then targetKey = Request("id") & ""
        If targetKey = "" Then targetKey = Request("idx") & ""
    End If

    ' 5. POST Form 인자 자동 수집 및 비밀번호 마스킹 (멀티파트 요청인 경우 스킵)
    formParamStr = ""
    If reqMethodUpper = "POST" And Not isMultipart Then
        For Each itemKey In Request.Form
            itemVal = Request.Form(itemKey) & ""
            paramKeyLower = LCase(itemKey)

            ' 비밀번호 및 보안 인자 마스킹 처리
            If InStr(paramKeyLower, "pwd") > 0 Or InStr(paramKeyLower, "pass") > 0 Then
                itemVal = "*****"
            End If

            If itemVal <> "" And itemKey <> "__VIEWSTATE" Then
                If formParamStr <> "" Then formParamStr = formParamStr & ", "
                formParamStr = formParamStr & itemKey & "=" & Left(itemVal, 50)
            End If
        Next
    End If

    ' 6. GET QueryString 인자 자동 수집 및 마스킹
    queryParamStr = ""
    If Request.QueryString.Count > 0 Then
        For Each itemKey In Request.QueryString
            itemVal = Request.QueryString(itemKey) & ""
            paramKeyLower = LCase(itemKey)

            If InStr(paramKeyLower, "pwd") > 0 Or InStr(paramKeyLower, "pass") > 0 Then
                itemVal = "*****"
            End If

            If itemVal <> "" Then
                If queryParamStr <> "" Then queryParamStr = queryParamStr & ", "
                queryParamStr = queryParamStr & itemKey & "=" & Left(itemVal, 50)
            End If
        Next
    End If

    ' 7. 상세 내역(LogDesc) 자동 생성
    logDesc = "[" & actionKorName & "] " & scriptName

    If formParamStr <> "" Then
        logDesc = logDesc & " [파라미터: " & formParamStr & "]"
    ElseIf queryParamStr <> "" Then
        logDesc = logDesc & " [쿼리: " & queryParamStr & "]"
    End If

    ' 8. 감사 로그 저장
    Call AddAuditLog(logType, "", targetKey, logDesc)
End Sub

' 공통 설정 로드 시 자동 실행
Call AutoDetectAuditLog()
%>