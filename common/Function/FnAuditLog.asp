<%
'====================================================================================
' * Function : AddAuditLog
' * Description : 관리자 감사 로그 DB 저장 공통 함수 (메뉴명 자동 추출 지원)
'====================================================================================
Sub AddAuditLog(argLogType, argMenuCode, argTargetKey, argLogDesc)
    On Error Resume Next

    Dim adminSeq, adminID, adminName, userIP, sqlQuery
    Dim menuCode, menuName, scriptName

    scriptName = LCase(Request.ServerVariables("SCRIPT_NAME")) ' 현재 요청 URL (예: /admin/member/mem_proc.asp)
    adminSeq  = Session("ASeq")
    adminID   = Session("AID")
    adminName = Session("AName")

    ' 1차 로그인 성공 후 OTP 인증 단계(임시 세션 상태)인 경우, 임시 세션 정보를 감사 로그에 반영
    If (IsNull(adminSeq) Or adminSeq = "") And Session("TempASeq") <> "" Then
        adminSeq  = Session("TempASeq")
        adminID   = Session("TempAID")
        adminName = Session("TempAName")
    End If

    If IsNull(adminSeq) Or adminSeq = "" Then adminSeq = 0
    If IsNull(adminID) Or adminID = "" Then adminID = "SYSTEM"
    If IsNull(adminName) Or adminName = "" Then adminName = "미인증/시스템"

    ' Client IP 획득 (프록시/로드밸런서 대응)
    userIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If userIP = "" Then userIP = Request.ServerVariables("REMOTE_ADDR")
    If InStr(userIP, ",") > 0 Then userIP = Trim(Split(userIP, ",")(0))

    ' 1. Class.FileMap (localMap 객체)을 활용한 메뉴코드 및 메뉴명 100% 자동 매핑
    Dim localMap, isLocalMapCreated
    isLocalMapCreated = False
    
    If IsObject(clsMap) Then
        If Not (clsMap Is Nothing) Then
            Set localMap = clsMap
        End If
    End If
    
    If localMap Is Nothing Then
        Set localMap = New ClsFileMap
        localMap.setInitDicAdmin()
        isLocalMapCreated = True
    End If

    If argMenuCode <> "" Then
        menuCode = argMenuCode
    Else
        menuCode = localMap.getMenuCode(scriptName)
    End If

    ' URL 경로 기반 메뉴명 자동 추출
    menuName = localMap.getMenuName(scriptName)



    ' [디버그] 브라우저 콘솔로그 출력을 통한 연동 결과 확인 코드 추가 (proc, ajax, download 페이지 제외 및 Null 에러 방지)
    ' 1. 변수 초기화 (Null 방지 및 VBScript 변수명 규칙 준수)
    Dim tmpScriptName, tmpMenuCode, tmpMenuName, tmpSearchName
    tmpScriptName = scriptName & ""
    tmpMenuCode = menuCode & ""
    tmpMenuName = menuName & ""
    tmpSearchName = LCase(tmpScriptName)
    
    ' 2. 대상 페이지 필터링
    If InStr(tmpSearchName, "_proc") = 0 And InStr(tmpSearchName, "_ajax") = 0 And InStr(tmpSearchName, "down") = 0 Then
        ' 3. 특수문자 이스케이프 처리
        tmpScriptName = Replace(tmpScriptName, Chr(39), Chr(92) & Chr(39))
        tmpMenuCode = Replace(tmpMenuCode, Chr(39), Chr(92) & Chr(39))
        tmpMenuName = Replace(tmpMenuName, Chr(39), Chr(92) & Chr(39))
        
        ' 4. 콘솔 출력
        'Response.Write "<script>console.log('AddAuditLog Debug - scriptName: " & tmpScriptName & ", menuCode: " & tmpMenuCode & ", menuName: " & tmpMenuName & "');</script>"
    End If
    
    If menuName = "" Or menuName = "메인" Then
        ' 경로 폴더 기반 기본 메뉴명 보완 매핑
        Select Case True
            Case InStr(scriptName, "/member/") > 0   : menuName = "회원관리"
            Case InStr(scriptName, "/recruit/") > 0  : menuName = "채용관리"
            Case InStr(scriptName, "/board/") > 0    : menuName = "게시글관리"
            Case InStr(scriptName, "/project/") > 0  : menuName = "프로젝트관리"
            Case InStr(scriptName, "/siteconf/") > 0 : menuName = "환경설정"
            Case InStr(scriptName, "/stats/") > 0    : menuName = "통계관리"
            Case Else : menuName = "관리자시스템"
        End Select
    End If

    ' 대메뉴명 (ParentMenuName) 추출 로직
    Dim parentMenuName
    parentMenuName = ""
    Select Case Left(menuCode, 2)
        Case "01" : parentMenuName = "채용관리"
        Case "02" : parentMenuName = "회원관리"
        Case "03" : parentMenuName = "프로젝트관리"
        Case "04" : parentMenuName = "게시글관리"
        Case "05" : parentMenuName = "통계관리"
        Case "06" : parentMenuName = "환경설정"
        Case "07" : parentMenuName = "쪽지함"
        Case Else
            ' 메뉴 코드가 없거나 매핑되지 않은 경우 폴더 경로 기반 판별 (레퍼러 보완 적용)
            Dim tmpReferer : tmpReferer = Request.ServerVariables("HTTP_REFERER") & ""
            Dim checkPath : checkPath = scriptName
            If InStr(checkPath, "/common/") > 0 And tmpReferer <> "" Then
                Dim tmpPos
                tmpPos = InStr(tmpReferer, "://")
                If tmpPos > 0 Then tmpReferer = Mid(tmpReferer, tmpPos + 3)
                tmpPos = InStr(tmpReferer, "/")
                If tmpPos > 0 Then checkPath = Mid(tmpReferer, tmpPos)
                tmpPos = InStr(checkPath, "?")
                If tmpPos > 0 Then checkPath = Left(checkPath, tmpPos - 1)
            End If
            
            Select Case True
                Case InStr(checkPath, "/recruit/") > 0  : parentMenuName = "채용관리"
                Case InStr(checkPath, "/member/") > 0   : parentMenuName = "회원관리"
                Case InStr(checkPath, "/project/") > 0  : parentMenuName = "프로젝트관리"
                Case InStr(checkPath, "/board/") > 0    : parentMenuName = "게시글관리"
                Case InStr(checkPath, "/stats/") > 0    : parentMenuName = "통계관리"
                Case InStr(checkPath, "/siteconf/") > 0 : parentMenuName = "환경설정"
                Case InStr(checkPath, "/message/") > 0  : parentMenuName = "쪽지함"
                Case Else : parentMenuName = "관리자시스템"
            End Select
    End Select

    ' 서브메뉴명 자동 판별 및 특정 페이지 재매핑 (배열 기반 검색)
    Dim subMenuName, arrStatsPages, statsPageItem, isStatsPage
    arrStatsPages = Array("job_partner_stats", _
                          "job_apply_stats", _
                          "job_apply_user_xls_new", _
                          "job_recom_list", _
                          "job_apply_user", _
                          "job_apply_user_xls", _
                          "job_apply_proc", _
                          "filedown")
    isStatsPage = False
    
    For Each statsPageItem In arrStatsPages
        If InStr(scriptName, statsPageItem) > 0 Then
            isStatsPage = True
            Exit For
        End If
    Next

    If isStatsPage Then
        Dim refererUrl, refererPath, pos, parentMenuNameFromRef
        refererUrl = Request.ServerVariables("HTTP_REFERER") & ""
        
        ' 1단계: 레퍼러 URL 파싱
        If refererUrl <> "" Then
            pos = InStr(refererUrl, "://")
            If pos > 0 Then refererUrl = Mid(refererUrl, pos + 3)
            
            pos = InStr(refererUrl, "/")
            If pos > 0 Then
                refererPath = Mid(refererUrl, pos)
            Else
                refererPath = ""
            End If
            
            pos = InStr(refererPath, "?")
            If pos > 0 Then refererPath = Left(refererPath, pos - 1)
        End If
        
        parentMenuNameFromRef = ""
        
        ' 2단계: 만약 레퍼러가 자기 자신이 아니라면(최초 팝업 진입), 레퍼러 파싱을 최우선으로 적용!
        Dim isSelfReferer : isSelfReferer = False
        If refererPath <> "" Then
            Dim currentFile, refererFile
            currentFile = LCase(Mid(scriptName, InStrRev(scriptName, "/") + 1) & "")
            refererFile = LCase(Mid(refererPath, InStrRev(refererPath, "/") + 1) & "")
            
            ' 서브 액션 페이지들(xls_new를 제외한 xls, proc, filedown)이 실행 중이거나 직전 레퍼러인 경우 모두 동일 컨텍스트로 취급(isSelfReferer = True)
            Dim isCurrentSubAction : isCurrentSubAction = (InStr(currentFile, "xls") > 0 And InStr(currentFile, "xls_new") = 0) Or InStr(currentFile, "proc") > 0 Or InStr(currentFile, "filedown") > 0
            Dim isRefererSubAction : isRefererSubAction = (InStr(refererFile, "xls") > 0 And InStr(refererFile, "xls_new") = 0) Or InStr(refererFile, "proc") > 0 Or InStr(refererFile, "filedown") > 0
            
            If isCurrentSubAction Or isRefererSubAction Then
                isSelfReferer = True
            Else
                currentFile = Replace(currentFile, ".asp", "")
                refererFile = Replace(refererFile, ".asp", "")
                If InStr(currentFile, refererFile) > 0 Or InStr(refererFile, currentFile) > 0 Then
                    isSelfReferer = True
                End If
            End If
        End If
        
        If Not isSelfReferer And refererPath <> "" Then
            Err.Clear
            parentMenuNameFromRef = localMap.getMenuName(LCase(refererPath))
            Err.Clear
            
            ' 최초 팝업 진입 시 성공 감지한 부모 대메뉴명을 후속 액션(엑셀, 다운로드 등)에서 꺼내 쓸 수 있도록 세션에 실시간 백업 보관!
            If parentMenuNameFromRef <> "" And parentMenuNameFromRef <> "메인" Then
                Session("ParentMenuName") = parentMenuNameFromRef
            End If
        End If
        
        ' 3단계: 만약 최초 레퍼러 분석에 실패했거나 자기 자신을 호출한 재요청(isSelfReferer = True)인 경우, 파라미터/세션을 순차적으로 백업 조회!
        If parentMenuNameFromRef = "" Or parentMenuNameFromRef = "메인" Then
            parentMenuNameFromRef = GetAuditParam("ParentMenuName") & ""
            If parentMenuNameFromRef = "" Then
                parentMenuNameFromRef = Request.Cookies("ParentMenuName") & ""
            End If
            If parentMenuNameFromRef = "" Then
                parentMenuNameFromRef = Session("ParentMenuName") & ""
            End If
        End If
        
        ' 파라미터/쿠키에 인코딩된 상태(% 문자 포함)로 들어온 경우 일괄적으로 순수 한글로 디코딩 보정
        If InStr(parentMenuNameFromRef, "%") > 0 Then
            parentMenuNameFromRef = Local_URLDecode(parentMenuNameFromRef)
        End If
        
        ' 4단계: 여전히 메뉴명이 없다면, RecrSeq를 활용한 DB 직접 조회 백업 처리!
        If parentMenuNameFromRef = "" Or parentMenuNameFromRef = "메인" Then
            Dim targetRecrSeq : targetRecrSeq = Trim(GetAuditParam("RecrSeq") & "")
            
            ' RecrSeq가 누락되고 Seq 파라미터가 단일 숫자(체크박스 목록이 아님)로 유입된 경우 공고번호로 차용
            If targetRecrSeq = "" Then
                Dim tmpSeqVal : tmpSeqVal = Trim(GetAuditParam("Seq") & "")
                If IsNumeric(tmpSeqVal) And InStr(tmpSeqVal, ",") = 0 Then
                    targetRecrSeq = tmpSeqVal
                End If
            End If
            
            ' 파일 다운로드(filedown) 시 파라미터 유실에 대한 UsrSeq/UsrID 기반 공고번호(RecrSeq) 역추적
            If targetRecrSeq = "" And InStr(scriptName, "filedown") > 0 Then
                Dim userSeqVal : userSeqVal = Trim(GetAuditParam("dnSeq") & "")
                If userSeqVal <> "" Then
                    Dim dbRs2, dbSql2, realUsrSeq
                    realUsrSeq = ""
                    
                    ' 숫자가 아닌 아이디 문자열이 넘어온 경우 TBL_USER에서 UsrSeq(숫자)를 먼저 조회
                    If Not IsNumeric(userSeqVal) Then
                        Dim dbRsID, dbSqlID
                        Err.Clear
                        Dim safeUsrID : safeUsrID = Replace(userSeqVal, "'", "''")
                        dbSqlID = "SELECT UsrSeq FROM TBL_USER WITH (nolock) WHERE UsrID = '" & safeUsrID & "'"
                        Set dbRsID = objDbCon.Execute(dbSqlID)
                        If Err.Number = 0 Then
                            If Not (dbRsID.Eof Or dbRsID.Bof) Then
                                realUsrSeq = dbRsID("UsrSeq") & ""
                            End If
                        End If
                        If IsObject(dbRsID) Then
                            If dbRsID.State = 1 Then dbRsID.Close
                            Set dbRsID = Nothing
                        End If
                    Else
                        realUsrSeq = userSeqVal
                    End If
                    
                    ' 획득한 UsrSeq(숫자)로 TBL_APPLY에서 채용공고 번호 역조회
                    If IsNumeric(realUsrSeq) And realUsrSeq <> "" Then
                        Err.Clear
                        dbSql2 = "SELECT TOP 1 RecrSeq FROM TBL_APPLY WITH (nolock) WHERE UsrSeq = " & realUsrSeq
                        Set dbRs2 = objDbCon.Execute(dbSql2)
                        If Err.Number = 0 Then
                            If Not (dbRs2.Eof Or dbRs2.Bof) Then
                                targetRecrSeq = dbRs2("RecrSeq") & ""
                            End If
                        End If
                        If IsObject(dbRs2) Then
                            If dbRs2.State = 1 Then dbRs2.Close
                            Set dbRs2 = Nothing
                        End If
                    End If
                    Err.Clear
                End If
            End If
            
            If IsNumeric(targetRecrSeq) And targetRecrSeq <> "" Then
                Dim dbRs, dbSql
                Err.Clear
                dbSql = "SELECT IsView FROM TBL_RECRUIT WITH (nolock) WHERE RecrSeq = " & targetRecrSeq
                Set dbRs = objDbCon.Execute(dbSql)
                If Err.Number = 0 Then
                    If Not (dbRs.Eof Or dbRs.Bof) Then
                        Dim isViewVal : isViewVal = Trim(dbRs("IsView") & "")
                        If isViewVal = "-1" Then
                            parentMenuNameFromRef = "전체 채용공고"
                        ElseIf isViewVal = "True" Or isViewVal = "1" Then
                            parentMenuNameFromRef = "진행중인 채용공고"
                        Else
                            parentMenuNameFromRef = "마감된 채용공고"
                        End If
                    End If
                End If
                If IsObject(dbRs) Then
                    If dbRs.State = 1 Then dbRs.Close
                    Set dbRs = Nothing
                End If
                Err.Clear
            End If
        End If
        
        ' 5단계: 그 외 예외적인 경우에 한하여 기존 레퍼러 매핑 2순위 재시도
        If (parentMenuNameFromRef = "" Or parentMenuNameFromRef = "메인") And refererPath <> "" Then
            Err.Clear
            parentMenuNameFromRef = localMap.getMenuName(LCase(refererPath))
                If Err.Number <> 0 Then
                    Dim errDesc
                    errDesc = CStr(Err.Description & "")
                    errDesc = Replace(errDesc, "'", "\'")
                    errDesc = Replace(errDesc, vbCrLf, " ")
                    errDesc = Replace(errDesc, vbCr, " ")
                    errDesc = Replace(errDesc, vbLf, " ")
                    'Response.Write "<script>console.log('Error in getMenuName: (" & Err.Number & ") " & errDesc & "');</script>"
                End If
            End If
            'Response.Write "<script>console.log('isLocalMapCreated: " & isLocalMapCreated & "');</script>"
            
            Dim safeRefererPath
            safeRefererPath = CStr(refererPath & "")
            safeRefererPath = Replace(safeRefererPath, "'", "\'")
            'Response.Write "<script>console.log('refererPath (LCase): " & LCase(safeRefererPath) & "');</script>"

            Dim safeParentMenuName
            safeParentMenuName = CStr(parentMenuNameFromRef & "")
            safeParentMenuName = Replace(safeParentMenuName, "'", "\'")
            safeParentMenuName = Replace(safeParentMenuName, vbCrLf, " ")
            safeParentMenuName = Replace(safeParentMenuName, vbCr, " ")
            safeParentMenuName = Replace(safeParentMenuName, vbLf, " ")
            'Response.Write "<script>console.log('parentMenuNameFromRef: " & safeParentMenuName & "');</script>"
        
        ' isStatsPage 이면서 내부 서브 액션(isSelfReferer = True)인 경우에 한해서만 subMenuName 을 레퍼러(부모 팝업) 주소 기준으로 복원!
        subMenuName = ""
        If isSelfReferer And refererPath <> "" Then
            Dim tmpSubMenu : tmpSubMenu = localMap.getMenuName(LCase(refererPath))
            If InStr(tmpSubMenu, vbTab) > 0 Then tmpSubMenu = Split(tmpSubMenu, vbTab)(0)
            If InStr(tmpSubMenu, "\t") > 0 Then tmpSubMenu = Split(tmpSubMenu, "\t")(0)
            subMenuName = Trim(tmpSubMenu & "")
        Else
            ' 최초 진입(isSelfReferer = False)일 때는 이미 구한 본래의 파일 메뉴명(menuName)을 그대로 소분류로 사용!
            subMenuName = menuName
        End If
        
        ' 레퍼러 유실 대안 폴더/파일명 및 전송 파라미터 기반 동적 매핑 (하드코딩 배제)
        If subMenuName = "" Or subMenuName = "관리자시스템" Then
            Dim reqFlag : reqFlag = UCase(Trim(GetAuditParam("dnFlag") & GetAuditParam("Flag") & ""))
            Select Case True
                Case InStr(scriptName, "stats") > 0 
                    subMenuName = "통계관리"
                Case InStr(scriptName, "down") > 0 Or InStr(scriptName, "xls") > 0
                    Select Case reqFlag
                        Case "USERS", "RECR" : subMenuName = "지원자관리"
                        Case "BBS"          : subMenuName = "게시판관리"
                        Case "CONF"         : subMenuName = "환경설정"
                        Case "MEMBER", "MEM" : subMenuName = "회원관리"
                        Case Else
                            If parentMenuName <> "" And parentMenuName <> "관리자시스템" Then
                                subMenuName = parentMenuName
                            Else
                                subMenuName = "지원자관리"
                            End If
                    End Select
                Case Else
                    If parentMenuName <> "" And parentMenuName <> "관리자시스템" Then
                        subMenuName = parentMenuName
                    Else
                        subMenuName = "지원자관리"
                    End If
            End Select
        End If
        
        If parentMenuNameFromRef <> "" And parentMenuNameFromRef <> "메인" Then
            menuName = parentMenuNameFromRef ' 예: "진행중인 채용공고"
            
            ' parentMenuNameFromRef 기준 동적 대메뉴명 복원 (하드코딩 방지)
            Dim dictKey, resolvedCode, resolvedParentPrefix
            resolvedCode = ""
            For Each dictKey In localMap.DicMenuCode.Keys
                If IsArray(localMap.DicMenuCode(dictKey)) Then
                    Dim curMenuNameVal
                    curMenuNameVal = localMap.DicMenuCode(dictKey)(1)
                    If InStr(curMenuNameVal, vbTab) > 0 Then curMenuNameVal = Split(curMenuNameVal, vbTab)(0)
                    If InStr(curMenuNameVal, "\t") > 0 Then curMenuNameVal = Split(curMenuNameVal, "\t")(0)
                    
                    If Trim(curMenuNameVal & "") = Trim(parentMenuNameFromRef & "") Then
                        resolvedCode = localMap.DicMenuCode(dictKey)(0)
                        Exit For
                    End If
                End If
            Next
            
            If resolvedCode <> "" Then
                resolvedParentPrefix = Left(resolvedCode, 2)
                Select Case resolvedParentPrefix
                    Case "01" : parentMenuName = "채용관리"
                    Case "02" : parentMenuName = "회원관리"
                    Case "03" : parentMenuName = "프로젝트관리"
                    Case "04" : parentMenuName = "게시글관리"
                    Case "05" : parentMenuName = "통계관리"
                    Case "06" : parentMenuName = "환경설정"
                      Case "07" : parentMenuName = "쪽지함"
                End Select
            End If
        Else
            menuName = "채용공고 관리" ' 매핑 실패 시 디폴트 기본값
        End If
    Else
        subMenuName = menuName
    End If

    If isLocalMapCreated Then
        Set localMap = Nothing
    End If

    ' SQL Injection 방지 특수문자 치환
    adminID        = Replace(adminID, "'", "''")
    adminName      = Replace(adminName, "'", "''")
    userIP         = Replace(userIP, "'", "''")
    menuCode       = Replace(menuCode, "'", "''")
    parentMenuName = Replace(parentMenuName, "'", "''")
    menuName       = Replace(menuName, "'", "''")
    subMenuName    = Replace(subMenuName, "'", "''")
    argLogDesc     = Replace(argLogDesc, "'", "''")

    ' 일반 INSERT 쿼리 실행 (대메뉴명 및 메뉴명 자동 기록)
    sqlQuery = "INSERT INTO TBL_ADMIN_AUDIT_LOG (" & _
               "    AdminSeq, AdminID, AdminName, AdminIP, MenuCode, ParentMenuName, MenuName, SubMenuName, LogType, TargetKey, LogDesc, RegDate" & _
               ") VALUES (" & _
               "     " & adminSeq & ", " & _
               "    '" & adminID & "', " & _
               "    '" & adminName & "', " & _
               "    '" & userIP & "', " & _
               "    '" & menuCode & "', " & _
               "    '" & parentMenuName & "', " & _
               "    '" & menuName & "', " & _
               "    '" & subMenuName & "', " & _
               "    '" & UCase(argLogType) & "', " & _
               "    '" & argTargetKey & "', " & _
               "    '" & argLogDesc & "', " & _
               "    GETDATE()" & _
               ")"

    ' 이전 처리 과정의 에러 객체를 초기화하여 실제 DB 실행 에러만 정확히 감지하도록 합니다.
    Err.Clear 
    objDbCon.Execute sqlQuery

    Dim debugMsg
    debugMsg = "AddAuditLog Console Debug -" & _
               " Script: " & (scriptName & "") & _
               ", Param: " & (GetAuditParam("ParentMenuName") & "") & _
               ", Session: " & (Session("ParentMenuName") & "") & _
               ", Referer: " & (Request.ServerVariables("HTTP_REFERER") & "") & _
               ", isStatsPage: " & (isStatsPage & "") & _
               ", parentMenuNameFromRef: " & (parentMenuNameFromRef & "") & _
               ", Final menuName: " & (menuName & "") & _
               ", Final subMenuName: " & (subMenuName & "")
    'Response.Write "<script>console.log(`" & debugMsg & "`);</script>"
End Sub

Function GetAuditParam(paramName)
    Dim isMultipart : isMultipart = (InStr(LCase(Request.ServerVariables("CONTENT_TYPE")), "multipart/form-data") > 0)
    If isMultipart Then
        GetAuditParam = Request.QueryString(paramName)
    Else
        GetAuditParam = Request(paramName)
    End If
End Function
%>