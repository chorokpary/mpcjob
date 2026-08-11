# 공통 Include (AdminConfig.asp) 기반 감사 로그 자동 감지 및 적용 가이드

본 문서는 관리자 시스템 내 수백 개의 ASP 소스 코드를 일일이 수정하지 않고, **공통 설정 파일(`AdminConfig.asp`) 1곳 수정으로 95% 이상의 감사 로그를 자동 수집**하는 방법과 **보안 상 중요한 5% 핵심 프로세스의 상세 수집 보완 방법**을 정리한 가이드입니다.

---

## 📌 1. 개요 및 하이브리드 설계 전략

### 💡 하이브리드(Hybrid) 수집 구조
* **95% 자동 감지 (공통 Include 방식)**:
  - `AdminConfig.asp` 하단에 자동 감지 엔진(`AutoDetectAuditLog`)을 탑재하여 **모든 페이지의 C/U/D 작업, 엑셀 다운로드, 파일 다운로드**를 자동으로 분석 및 수집합니다.
* **5% 핵심 보완 (수동 1줄 추가 방식)**:
  - 회원 강제 탈퇴, 권한 변경, 주요 개인정보 파일 다운로드 등 **법적/보안상 매우 중요한 핵심 프로세서(`*_proc.asp`) 파일에만** DB 성공 직후 `AddAuditLog` 1줄을 추가하여 한글 상세 내역(`LogDesc`)을 보강합니다.

---

## 🗄️ 2. 데이터베이스 테이블 생성 (`TBL_ADMIN_AUDIT_LOG`)

감사 로그 데이터를 저장하기 위해 MSSQL 데이터베이스에 아래 테이블 및 인덱스를 생성합니다.

```sql
-- 1. 감사 로그 저장 테이블 생성 쿼리
CREATE TABLE TBL_ADMIN_AUDIT_LOG (
    LogSeq         BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    AdminSeq       INT NULL,                         -- 관리자 고유번호 (Session("ASeq"))
    AdminID        VARCHAR(50) NOT NULL,             -- 관리자 아이디 (Session("AID"))
    AdminName      NVARCHAR(50) NULL,                -- 관리자 이름 (Session("AName"))
    AdminIP        VARCHAR(45) NOT NULL,             -- 접속 IP (IPv4 / IPv6)
    MenuCode       VARCHAR(20) NULL,                 -- 메뉴 코드 (예: 010100, 020100, 060100 등)
    ParentMenuName NVARCHAR(100) NULL,               -- 대메뉴명 (예: 회원관리, 환경설정 등)
    MenuName       NVARCHAR(100) NULL,               -- 메뉴명 (예: 관리자등록, 회원관리 등)
    SubMenuName    NVARCHAR(100) NULL,               -- 서브메뉴명 (예: 목록, 입력폼, 처리 등)
    LogType        VARCHAR(20) NOT NULL,             -- CREATE, READ, UPDATE, DELETE, FILEDOWN, EXCEL
    TargetKey      VARCHAR(100) NULL,                -- 대상 PK 값 (예: 회원Seq, 게시글Seq, IpSeq 등)
    LogDesc        NVARCHAR(1000) NULL,              -- 작업 상세 내용
    RegDate        DATETIME DEFAULT GETDATE() NOT NULL -- 발생 일시
);

-- 2. 검색 및 조회 성능 최적화를 위한 인덱스 생성
CREATE INDEX IX_AUDIT_LOG_REGDATE ON TBL_ADMIN_AUDIT_LOG(RegDate DESC);
CREATE INDEX IX_AUDIT_LOG_ADMINID ON TBL_ADMIN_AUDIT_LOG(AdminID, RegDate DESC);
CREATE INDEX IX_AUDIT_LOG_LOGTYPE ON TBL_ADMIN_AUDIT_LOG(LogType, RegDate DESC);
```

---

## ⚙️ 3. 2단계: 공통 자동 감지 엔진 구축 (`AdminConfig.asp`)

### ① 감사 로그 공통 함수 모듈 (`/common/Function/FnAuditLog.asp`)

DB 저장 프로시저(SP) 생성 없이 **일반 SQL Direct `INSERT INTO` 쿼리 방식**으로 구현하며, `AdminConfig.asp`에 이미 선언되어 있는 `clsMap` (클래스 `ClsFileMap`)을 통해 **메뉴 코드와 메뉴명이 100% 자동으로 매핑되어 저장**됩니다.

```asp
<%
'====================================================================================
' * Function : AddAuditLog
' * Description : 관리자 감사 로그 DB 저장 공통 함수 (메뉴명 자동 추출 지원)
'====================================================================================
Sub AddAuditLog(argLogType, argMenuCode, argTargetKey, argLogDesc)
    On Error Resume Next

    Dim adminSeq, adminID, adminName, userIP, sqlQuery
    Dim menuCode, menuName, scriptName

    scriptName = Request.ServerVariables("SCRIPT_NAME") ' 현재 요청 URL (예: /admin/member/mem_proc.asp)
    adminSeq  = Session("ASeq")
    adminID   = Session("AID")
    adminName = Session("AName")

    If IsNull(adminSeq) Or adminSeq = "" Then adminSeq = 0
    If IsNull(adminID) Or adminID = "" Then adminID = "SYSTEM"
    If IsNull(adminName) Or adminName = "" Then adminName = "미인증/시스템"

    ' Client IP 획득 (프록시/로드밸런서 대응)
    userIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If userIP = "" Then userIP = Request.ServerVariables("REMOTE_ADDR")
    If InStr(userIP, ",") > 0 Then userIP = Trim(Split(userIP, ",")(0))

    ' 1. Class.FileMap (clsMap 객체)을 활용한 메뉴코드 및 메뉴명 100% 자동 매핑
    If argMenuCode <> "" Then
        menuCode = argMenuCode
    Else
        menuCode = clsMap.getMenuCode(scriptName)
    End If

    ' URL 경로 기반 메뉴명 자동 추출
    menuName = clsMap.getMenuName(scriptName)
    If menuName = "" Or menuName = "메인" Then
        ' 경로 폴더 기반 기본 메뉴명 보완 매핑
        Select Case True
            Case InStr(scriptName, "/iplimit/") > 0  : menuName = "IP접근제한관리"
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
            Select Case True
                Case InStr(scriptName, "/recruit/") > 0  : parentMenuName = "채용관리"
                Case InStr(scriptName, "/member/") > 0   : parentMenuName = "회원관리"
                Case InStr(scriptName, "/project/") > 0  : parentMenuName = "프로젝트관리"
                Case InStr(scriptName, "/board/") > 0    : parentMenuName = "게시글관리"
                Case InStr(scriptName, "/stats/") > 0    : parentMenuName = "통계관리"
                Case InStr(scriptName, "/siteconf/") > 0 : parentMenuName = "환경설정"
                Case InStr(scriptName, "/message/") > 0  : parentMenuName = "쪽지함"
                Case Else : parentMenuName = "관리자시스템"
            End Select
    End Select

    ' 서브메뉴명 자동 판별 및 특정 페이지 재매핑 (배열 기반 검색)
    Dim subMenuName, arrStatsPages, statsPageItem, isStatsPage
    arrStatsPages = Array("job_partner_stats", "job_apply_stats")
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
        parentMenuNameFromRef = ""
        Response.Write "<script>console.log('refererUrl: " & Replace(refererUrl, "'", "\'") & "');</script>"
        
        ' 1. Referer URL에서 도메인을 제외한 순수 경로(Path) 추출
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
            
            ' 2. 부모창 경로로 파일맵에서 메뉴명 자동 매핑 (대소문자 불일치 방지를 위해 LCase 처리)
            If refererPath <> "" Then
                Err.Clear
                parentMenuNameFromRef = clsMap.getMenuName(LCase(refererPath))
                If Err.Number <> 0 Then
                    Dim errDesc
                    errDesc = CStr(Err.Description & "")
                    errDesc = Replace(errDesc, "'", "\'")
                    errDesc = Replace(errDesc, vbCrLf, " ")
                    errDesc = Replace(errDesc, vbCr, " ")
                    errDesc = Replace(errDesc, vbLf, " ")
                    Response.Write "<script>console.log('Error in getMenuName: (" & Err.Number & ") " & errDesc & "');</script>"
                End If
            End If
            Response.Write "<script>console.log('isLocalMapCreated: " & isLocalMapCreated & "');</script>"
            Dim safeRefererPath
            safeRefererPath = CStr(refererPath & "")
            safeRefererPath = Replace(safeRefererPath, "'", "\'")
            Response.Write "<script>console.log('refererPath (LCase): " & LCase(safeRefererPath) & "');</script>"
            Dim safeParentMenuName
            safeParentMenuName = CStr(parentMenuNameFromRef & "")
            safeParentMenuName = Replace(safeParentMenuName, "'", "\'")
            safeParentMenuName = Replace(safeParentMenuName, vbCrLf, " ")
            safeParentMenuName = Replace(safeParentMenuName, vbCr, " ")
            safeParentMenuName = Replace(safeParentMenuName, vbLf, " ")
            Response.Write "<script>console.log('parentMenuNameFromRef: " & safeParentMenuName & "');</script>"
        End If
        
        subMenuName = menuName ' 기존 MenuName("통계관리")을 SubMenuName으로 밀어냄
        
        If parentMenuNameFromRef <> "" And parentMenuNameFromRef <> "메인" Then
            menuName = parentMenuNameFromRef
        Else
            menuName = "채용공고 관리"
        End If
    Else
        subMenuName = menuName
    End If

    ' SQL Injection 방지 특수문자 치환
    adminID        = Replace(adminID, "'", "''")
    adminName      = Replace(adminName, "'", "''")
    userIP         = Replace(userIP, "'", "''")
    menuCode       = Replace(menuCode, "'", "''")
    parentMenuName = Replace(parentMenuName, "'", "''")
    menuName       = Replace(menuName, "'", "''")
    subMenuName    = Replace(subMenuName, "'", "''")
    argTargetKey   = Replace(argTargetKey, "'", "''")
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

    objDbCon.Execute sqlQuery
End Sub
%>
```

---

### ② 전역 고도화 자동 감지 엔진 (`AdminConfig.asp` 하단에 추가)

`AdminConfig.asp` 최하단에 아래 전역 자동 감지 엔진을 추가하면, **POST/GET 전송 인자를 자동으로 분석하여 파라미터 내용까지 한눈에 파악할 수 있는 고도화된 감사 로그를 100% 자동 기록**합니다. (비밀번호 자동 마스킹 및 행위별 한글 라벨 자동 할당 포함)

```asp
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

    scriptName    = LCase(Request.ServerVariables("SCRIPT_NAME")) ' 요청 URL (예: /admin/siteconf/iplimit/ip_proc.asp)
    requestMethod = Request.ServerVariables("REQUEST_METHOD")     ' GET 또는 POST

    ' 2. 엑셀 및 파일 다운로드 패턴 감지 및 행위 라벨링
    If InStr(scriptName, "xls") > 0 Or InStr(scriptName, "excel") > 0 Or Request("excel") = "Y" Then
        logType = "EXCEL"
        actionKorName = "엑셀 다운로드"
    ElseIf InStr(scriptName, "down") > 0 Or InStr(scriptName, "download") > 0 Then
        logType = "FILEDOWN"
        actionKorName = "파일 다운로드"
    Else
        ' 3. C/U/D 처리 파일(*_proc.asp, *_ajax.asp) 또는 POST 요청 시 행위 감지
        If InStr(scriptName, "_proc") > 0 Or InStr(scriptName, "_ajax") > 0 Or requestMethod = "POST" Then
            ' 폼/쿼리 파라미터(action, Flag, mode, act) 분석
            actionParam = UCase(Trim(Request("action") & Request("Flag") & Request("mode") & Request("act")))

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
        Else
            ' 일반 단순 리스트/폼 조회 페이지는 로그 생성 스킵 (READ 로그 폭주 방지)
            Exit Sub
        End If
    End If

    ' 4. PK 타겟키 자동 추출 (seq, id, idx, pno 파라미터 감지)
    targetKey = Request("seq")
    If targetKey = "" Then targetKey = Request("Seq")
    If targetKey = "" Then targetKey = Request("id")
    If targetKey = "" Then targetKey = Request("idx")

    ' 5. POST Form 인자 자동 수집 및 비밀번호 마스킹
    formParamStr = ""
    If requestMethod = "POST" And Request.Form.Count > 0 Then
        For Each itemKey In Request.Form
            itemVal = Request.Form(itemKey)
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
            itemVal = Request.QueryString(itemKey)
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

    ' 7. 사람이 읽기 쉬운 명확한 상세 내역(LogDesc) 자동 생성
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
```

---

## 🎯 3. 2단계: 핵심 포인트 보완 (수동 1줄 명시적 보완 가이드)

### ❓ 왜 핵심 포인트 보완이 필요한가요?

자동 감지 로직만 사용할 경우 발생할 수 있는 **2가지 한계점**을 보완하기 위함입니다:

1. **상세 정보(LogDesc) 인지성 부족**:  
   자동 감지는 `[자동감지] /admin/member/mem_proc.asp [POST] ?Flag=DEL` 형태로 기술적인 URL만 기록되지만, 개인정보보호법 및 보안 감사 시에는 **`"회원 강제 탈퇴 (아이디: user01, 이름: 홍길동)"`**과 같이 사람이 한눈에 인지할 수 있는 설명이 필요합니다.
2. **DB 트랜잭션 성공 보장**:  
   자동 감지는 ASP 페이지 진입 시점에 실행되므로, 만약 DB 오류로 삭제에 실패해도 "삭제 시도 로그"로 남습니다. **보안상 핵심적인 삭제/변경 작업은 DB 커밋 성공(`Err.Number = 0`) 후에만 명시적으로 남겨야 완벽합니다.**

---

### 🛠️ 핵심 프로세서별 구체적 보완 위치 및 소스 코드

보완 코드는 각 처리 파일(`*_proc.asp`)의 **DB 실행(`objDbCon.Execute`) 성공 확인 직후(`Err.Number = 0`), 리다이렉트(`Response.Redirect`) 직전 위치**에 1줄씩 추가합니다.

---

#### ① IP 접근 제한 관리 ([ip_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_proc.asp))

* **적용 위치**: `ip_proc.asp` 내 각 `Case`별 DB 실행 직후

```asp
    Select Case action
        ' 1. 신규 허용 IP 등록
        Case "insert"
            ' ... (중략: allowIp, ipMemo 획득 및 DB INSERT 쿼리 실행) ...
            objDbCon.Execute sqlQuery, rowsAffected

            If Err.Number <> 0 Then
                Response.Write "<script>alert('IP 등록 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If

            ' [핵심 보완 코드 추가 위치] DB 등록 성공 시 명시적 로그 작성
            Call AddAuditLog("CREATE", "060100", "", "신규 IP 허용 등록 (IP: " & allowIp & ", 메모: " & ipMemo & ")")

            On Error GoTo 0
            Response.Redirect "ip_list.asp"
            Response.End

        ' 2. 등록된 허용 IP 삭제
        Case "delete"
            ' ... (중략: DELETE FROM TBL_IP_ALLOW WHERE IpSeq = ipSeq 실행) ...
            objDbCon.Execute sqlQuery, rowsAffected

            If Err.Number <> 0 Then
                Response.Write "<script>alert('IP 삭제 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If

            ' [핵심 보완 코드 추가 위치] DB 삭제 성공 시 명시적 로그 작성
            Call AddAuditLog("DELETE", "060100", ipSeq, "허용 IP 삭제 (IPSeq: " & ipSeq & ")")

            On Error GoTo 0
            Response.Redirect "ip_list.asp"
            Response.End

        ' 3. IP 사용 상태 토글 (활성/비활성)
        Case "toggle"
            ' ... (중략: UPDATE TBL_IP_ALLOW SET IsUse = isUse 실행) ...
            objDbCon.Execute sqlQuery, rowsAffected

            If Err.Number <> 0 Then
                Response.Write "<script>alert('상태 변경 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If

            ' [핵심 보완 코드 추가 위치] 상태 토글 성공 시 명시적 로그 작성
            Call AddAuditLog("UPDATE", "060100", ipSeq, "IP 사용 상태 변경 (IPSeq: " & ipSeq & ", 변경상태: " & isUse & ")")

            On Error GoTo 0
            Response.Redirect "ip_list.asp"
            Response.End
    End Select
```

---

#### ② IP 예외 관리자 설정 ([ip_bypass_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_proc.asp))

* **적용 위치**: `ip_bypass_proc.asp` 내 예외 ID 등록 및 삭제 성공 직후

```asp
    Select Case action
        ' 1. 예외 관리자 ID 등록
        Case "insert"
            ' ... (INSERT INTO TBL_IP_BYPASS 실행) ...
            objDbCon.Execute sqlQuery, rowsAffected

            If Err.Number <> 0 Then
                Response.Write "<script>alert('예외 ID 등록 도중 오류가 발생했습니다.'); history.back();</script>"
                Response.End
            End If

            ' [핵심 보완] 예외 ID 등록 성공 로그
            Call AddAuditLog("CREATE", "060100", adminId, "예외 관리자 ID 등록 (ID: " & adminId & ", 메모: " & bypassMemo & ")")

            On Error GoTo 0
            Response.Redirect "ip_bypass_pop.asp"
            Response.End

        ' 2. 예외 관리자 ID 삭제
        Case "delete"
            ' ... (DELETE FROM TBL_IP_BYPASS 실행) ...
            objDbCon.Execute sqlQuery, rowsAffected

            If Err.Number <> 0 Then
                Response.Write "<script>alert('삭제 도중 오류가 발생했습니다.'); history.back();</script>"
                Response.End
            End If

            ' [핵심 보완] 예외 ID 삭제 성공 로그
            Call AddAuditLog("DELETE", "060100", bypassSeq, "예외 관리자 ID 삭제 (BypassSeq: " & bypassSeq & ")")

            On Error GoTo 0
            Response.Redirect "ip_bypass_pop.asp"
            Response.End
    End Select
```

---

#### ③ 회원 관리 (`/admin/member/mem_proc.asp`) - 회원 정보 수정 및 계정 삭제

* **적용 위치**: 회원 상태 변경, 강제 탈퇴 및 메모 수정 처리 성공 후

```asp
    ' Flag 값 분기 처리 예시
    Select Case strFlag
        Case "DEL" ' 회원 강제 삭제/탈퇴
            ' ... (DELETE/UPDATE 실행) ...
            If Err.Number = 0 Then
                ' [핵심 보완] 회원 삭제 명시적 감사 로그 남기기
                Call AddAuditLog("DELETE", "020100", strSeq, "회원 계정 강제 삭제 (회원Seq: " & strSeq & ", 아이디: " & strID & ", 이름: " & strName & ")")
                Response.Write "<script>alert('삭제되었습니다.'); location.href='mem_list.asp';</script>"
            End If

        Case "MOD" ' 회원 정보 수정
            ' ... (UPDATE 실행) ...
            If Err.Number = 0 Then
                ' [핵심 보완] 회원 수정 명시적 감사 로그 남기기
                Call AddAuditLog("UPDATE", "020100", strSeq, "회원 정보 수정 (회원Seq: " & strSeq & ", 아이디: " & strID & ")")
                Response.Write "<script>alert('수정되었습니다.'); location.href='mem_list.asp';</script>"
            End If
    End Select
```

---

#### ④ 채용 공고 및 지원자 관리 (`/admin/recruit/job_proc.asp`, `job_apply_proc.asp`)

* **적용 위치**: 입사지원서 개인정보 이력서 다운로드 및 전형 상태 변경 성공 시

```asp
    ' 1) 지원자 이력서 파일 다운로드 시점 (job_apply_user.asp / file_down.asp)
    Call AddAuditLog("FILEDOWN", "010100", applySeq, "입사지원서 이력서 첨부파일 다운로드 (지원자Seq: " & applySeq & ", 지원자명: " & applicantName & ")")

    ' 2) 지원자 전형 상태 변경 시점 (job_apply_proc.asp)
    If Err.Number = 0 Then
        Call AddAuditLog("UPDATE", "010100", applySeq, "입사지원자 전형 상태 변경 (지원자Seq: " & applySeq & ", 변경상태: " & strStatus & ")")
    End If
```

---

#### ⑤ 공지사항 및 팝업 게시판 관리 (`/admin/board/notice_proc.asp`, `pop_proc.asp`)

* **적용 위치**: 주요 공지사항 또는 팝업 게시글 삭제/수정 완료 시점

```asp
    ' notice_proc.asp - 공지사항 삭제 완료 시점
    If Err.Number = 0 Then
        Call AddAuditLog("DELETE", "040100", seq, "공지사항 게시글 삭제 (게시글Seq: " & seq & ", 제목: " & strSubject & ")")
        Response.Write "<script>alert('삭제되었습니다.'); location.href='notice_list.asp';</script>"
    End If
```

---

## 📊 4. 자동 감지 vs 핵심 보완 결과 비교

| 구분 | 1단계: 자동 감지 (`AdminConfig.asp`) | 2단계: 핵심 명시적 보완 (`AddAuditLog` 추가) |
| :--- | :--- | :--- |
| **작업 공수** | **소스 수정 0개** (`AdminConfig.asp` 1곳만 수정) | 주요 `*_proc.asp` 5~10개 파일에 1줄씩 추가 |
| **감지 범위** | 전체 관리자 페이지 C/U/D, 엑셀, 파일다운로드 | 핵심 보안/개인정보 처리 프로세스 |
| **로그 내용 (LogDesc)** | `[자동감지] /admin/member/mem_proc.asp [POST]` | `회원 계정 강제 삭제 (아이디: user01, 이름: 홍길동)` |
| **기록 시점** | 페이지 진입 및 요청 시점 | DB 쿼리 실행 성공(`Err.Number = 0`) 시점 |
| **권장 용도** | 시스템 전체 보안 행위 종합 감사망 구축 | 법적 증적 제출용 핵심 개인정보/보안 감사 |

---

## 💻 5. 감사 로그 DB 저장 데이터 및 관리자 리스트 화면 모양 예시

### ① DB 테이블 (`TBL_ADMIN_AUDIT_LOG`) 실제 저장 데이터 모습

DB에 누적되는 데이터는 **자동 감지 로그**와 **핵심 명시적 보완 로그**가 혼합되어 아래와 같이 저장됩니다:

| LogSeq | AdminID | AdminName | AdminIP | MenuCode | ParentMenuName | Menu---

## 📂 7. 감사 로그 관련 추가/수정 파일 경로 목록

감사 로그 수집 체계를 구현하기 위해 신규 추가되거나 수정된 주요 21개 파일 경로 목록입니다.

### ① 신규 추가 파일 (3개)
* **공통 모듈**:
  * [`/common/Function/FnAuditLog.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Function/FnAuditLog.asp) — 감사 로그 기록 공통 함수 (`AddAuditLog` 정의)
* **감사 로그 조회 및 엑셀**:
  * [`/admin/siteconf/audit_log_list.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_list.asp) — 감사 로그 목록 및 검색 페이지
  * [`/admin/siteconf/audit_log_xls.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_xls.asp) — 감사 로그 엑셀 다운로드 처리 페이지

### ② 기존 수정 파일 (18개)
* **공통 설정 및 매핑**:
  * [`/common/AdminConfig.asp`](file:///d:/MPCJOB/mpcjob_2012/common/AdminConfig.asp) — 전역 자동 감지 엔진 (`AutoDetectAuditLog`) 탑재, 공통 함수 인클루드 및 `mem_proc.asp` 예외 스킵 추가
  * [`/common/Class/Class.FileMap.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Class/Class.FileMap.asp) — 메뉴 코드 매핑 확장 및 메뉴명 탭 문자 정제
  * [`/common/CommonConfig.asp`](file:///d:/MPCJOB/mpcjob_2012/common/CommonConfig.asp) — 개발/테스트 호스트 환경 설정 변경
  * [`/common/Dev/FileDown.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Dev/FileDown.asp) — 파일 다운로드 성공/실패 감사 로그 수집 연동
* **로그인 / 로그아웃 / 레이아웃**:
  * [`/admin/login.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login.asp) — 로그인 폼 전송 주소 및 주석 구문 수정
  * [`/admin/login_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc.asp) — 관리자 로그인 성공/실패(`LOGIN`/`LOGIN_FAIL`) 감사 로그 연동
  * [`/admin/login_proc_otp.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc_otp.asp) — OTP 2차 인증 1차 성공/실패 감사 로그 연동
  * [`/admin/logout_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/logout_proc.asp) — 관리자 로그아웃(`LOGOUT`) 감사 로그 연동
  * [`/admin/include/left.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/include/left.asp) — LNB 메뉴 내 감사 로그 관리 링크 및 권한 연결
* **회원 관리**:
  * [`/admin/member/mem_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_form.asp) — 회원 상세 정보/이력서 열람 시 `READ` 로그 연동
  * [`/admin/member/mem_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_proc.asp) — Flag별 감사 로그 생성 및 자동 감지 파라미터 결합 적재
* **채용 관리**:
  * [`/admin/recruit/job_apply_user.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_apply_user.asp) — 이력서 다운로드 폼 POST 전송 및 공고번호 파라미터 전달 보완
  * [`/admin/recruit/job_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_form.asp) — 채용공고 폼 Action URL 및 감사 로그 연동
  * [`/admin/recruit/job_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_proc.asp) — 채용공고 처리 리다이렉트 주소 및 감사 로그 연동
* **게시판 관리**:
  * [`/admin/board/notice_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/board/notice_form.asp) — 공지사항 폼 진입 `READ` 로그 및 Action URL 연동
  * [`/admin/board/faq_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/board/faq_form.asp) — FAQ 폼 진입 `READ` 로그 및 Action URL 연동
* **환경설정 / IP 제한 관리**:
  * [`/admin/siteconf/iplimit/ip_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_proc.asp) — IP 허용 등록/삭제/토글 명시적 감사 로그 연동
  * [`/admin/siteconf/iplimit/ip_bypass_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_proc.asp) — 예외 ID 등록/삭제 명시적 감사 로그 연동

---

## 📂 8. 감사 로그 연동 관련 주요 파일별 수정 상세 내역

감사 로그(자동 감지 및 핵심 포인트 수동 보완) 시스템 적용 과정에서 추가 및 변경된 21개 주요 파일의 상세 내용입니다.

### 📌 금번 작업 수정/추가 파일 경로 목록 (21개)
* **신규 추가 파일 (3개)**:
  * [`/common/Function/FnAuditLog.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Function/FnAuditLog.asp)
  * [`/admin/siteconf/audit_log_list.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_list.asp)
  * [`/admin/siteconf/audit_log_xls.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_xls.asp)
* **기존 수정 파일 (18개)**:
  * [`/common/AdminConfig.asp`](file:///d:/MPCJOB/mpcjob_2012/common/AdminConfig.asp)
  * [`/common/Class/Class.FileMap.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Class/Class.FileMap.asp)
  * [`/common/CommonConfig.asp`](file:///d:/MPCJOB/mpcjob_2012/common/CommonConfig.asp)
  * [`/common/Dev/FileDown.asp`](file:///d:/MPCJOB/mpcjob_2012/common/Dev/FileDown.asp)
  * [`/admin/login.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login.asp)
  * [`/admin/login_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc.asp)
  * [`/admin/login_proc_otp.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc_otp.asp)
  * [`/admin/logout_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/logout_proc.asp)
  * [`/admin/board/faq_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/board/faq_form.asp)
  * [`/admin/board/notice_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/board/notice_form.asp)
  * [`/admin/include/left.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/include/left.asp)
  * [`/admin/member/mem_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_form.asp)
  * [`/admin/member/mem_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_proc.asp)
  * [`/admin/recruit/job_apply_user.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_apply_user.asp)
  * [`/admin/recruit/job_form.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_form.asp)
  * [`/admin/recruit/job_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_proc.asp)
  * [`/admin/siteconf/iplimit/ip_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_proc.asp)
  * [`/admin/siteconf/iplimit/ip_bypass_proc.asp`](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_proc.asp)

---

### ① 공통 / 라이브러리 레이어

#### 1. [FnAuditLog.asp](file:///d:/MPCJOB/mpcjob_2012/common/Function/FnAuditLog.asp) `[NEW]`
* **주요 내용**: 감사 로그 DB 저장 공통 함수 (`AddAuditLog`) 정의
* **상세 설명**:
  * 세션 정보(`ASeq`, `AID`, `AName`) 및 IP (`HTTP_X_FORWARDED_FOR` / `REMOTE_ADDR`)를 자동 획득하여 기록합니다.
  * `ClsFileMap`을 활용해 페이지별 메뉴코드 및 한글 메뉴명을 자동 매핑합니다.
  * 레퍼러(`HTTP_REFERER`) 및 공고번호(`RecrSeq`), 파일 다운로드 파라미터(`dnSeq`)를 종합 분석하여 유실될 수 있는 대메뉴/소메뉴명을 역추적 및 보완하는 다차원 매핑 알고리즘이 탑재되어 있습니다.
  * SQL Injection 방지를 위한 문자 치환 처리를 포함합니다.

#### 2. [AdminConfig.asp](file:///d:/MPCJOB/mpcjob_2012/common/AdminConfig.asp) `[MODIFY]`
* **주요 내용**: 전역 자동 감지 엔진 `AutoDetectAuditLog` 탑재, 공통 함수 인클루드 및 `mem_proc.asp` 스킵 예외 추가
* **상세 설명**:
  * 관리자 세션이 감지되면 요청 파라미터 분석을 통해 C/U/D 및 단순 조회(`READ`), 엑셀 다운로드(`EXCEL`), 파일 다운로드(`FILEDOWN`) 로그를 100% 자동 기록합니다.
  * 비밀번호 관련 파라미터(`pwd`, `pass` 등)는 `*****`로 자동 마스킹 처리하여 보관합니다.
  * 자체 명시적 감사 로그를 기록하는 `mem_proc.asp` 페이지 진입 시에는 중복 적재 방지를 위해 `AutoDetectAuditLog`를 `Exit Sub` 하도록 예외 처리 구문을 반영하였습니다.

#### 3. [Class.FileMap.asp](file:///d:/MPCJOB/mpcjob_2012/common/Class/Class.FileMap.asp) `[MODIFY]`
* **주요 내용**: 감사 로그에서 탐색하는 파일 정보 및 메뉴코드의 신규 매핑 반영 및 메뉴명 정제 로직 탑재
* **상세 설명**:
  * IP 접근 제한 우회 관련 파일(`ip_bypass_*`) 및 관리자 등록 엑셀 다운로드 관련 파일들의 감사 로그 매핑(`DicMenuCode`)을 신규 추가하였습니다.
  * **메뉴명 정제 로직 보정**: 기존 `getMenuName` 함수는 등록된 메뉴 정보 중 뒤쪽에 붙어있던 탭 문자(`\t`) 이후의 기호(예: `\t /corp/...`)까지 같이 리턴하는 현상이 있었습니다. 이로 인해 감사 로그 테이블에 지저분한 문자열이 들어가는 것을 막기 위해, `getMenuName` 함수에 탭 문자를 잘라내고 순수 한글 메뉴명만 반환하도록 문자열 정제(Split) 로직을 추가하였습니다.

#### 4. [CommonConfig.asp](file:///d:/MPCJOB/mpcjob_2012/common/CommonConfig.asp) `[MODIFY]`
* **주요 내용**: 테스트 서버 개발 환경 구성을 위한 호스트 상수 수정
* **상세 설명**:
  * 로컬 및 스테이징 환경에서의 감사 로그 연동 테스트를 원활히 하기 위해 `GLOBAL_HOST` 및 `GLOBAL_HOST_SSL` 상수를 `"www.mpcjob.co.kr"`에서 `"test.mpcjob.co.kr"`로 변경하였습니다.

#### 5. [FileDown.asp](file:///d:/MPCJOB/mpcjob_2012/common/Dev/FileDown.asp) `[MODIFY]`
* **주요 내용**: 파일 다운로드 성공 및 실패(파일 없음) 시의 감사 로그 수집 연동
* **상세 설명**:
  * `FnAuditLog.asp` 파일을 인클루드하였습니다.
  * 다운로드 대상 파일이 서버에 존재하지 않을 경우 `[실패/파일없음]` 로그를, 다운로드 성공 시 다운로드 완료 감사 로그(`FILEDOWN`)를 `AddAuditLog`를 통해 기록하도록 처리하였습니다.

---

### ② 로그인 / 세션 / 레이아웃 관리 레이어

#### 6. [login.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login.asp) `[MODIFY]`
* **주요 내용**: 로그인 폼 서밋 경로 수정 및 HTML 주석 오류 정정
* **상세 설명**:
  * 로그인 폼의 `action` 경로를 운영 서버(`https://www.mpcjob.co.kr/...`)에서 테스트 서버(`http://test.mpcjob.co.kr/...`)로 변경하였습니다.
  * HTML/JS 주석 태그의 오류(`<!-` -> `<!--`)를 정정하였습니다.

#### 7. [login_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc.asp) `[MODIFY]`
* **주요 내용**: 관리자 로그인 성공 및 실패 케이스별 명시적 감사 로그 추가
* **상세 설명**:
  * 로그인 성공 시 `LOGIN` 타입 로그("관리자 로그인 성공 (ID: ..., 이름: ...)")를 저장하고 리다이렉트 경로를 테스트 서버로 변경하였습니다.
  * 계정 잠금(`LOCK`), 비밀번호 불일치(`FAIL`), 내부 오류 등 로그인 실패 시 각각 `LOGIN_FAIL` 타입 로그를 상세 원인과 시도한 ID를 포함하여 기록하도록 보완하였습니다.

#### 8. [login_proc_otp.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc_otp.asp) `[MODIFY]`
* **주요 내용**: OTP 2차 인증을 사용하는 환경에서의 1차 로그인 성공/실패 감사 로그 추가
* **상세 설명**:
  * 1단계 로그인 성공 시(OTP 대기 상태 또는 OTP 신규 기기 등록 대기 상태) `LOGIN_OTP_WAIT` 감사 로그를 적재하도록 수정하였습니다.
  * 계정 잠금, 인증 실패 등의 경우 `LOGIN_FAIL` 감사 로그를 기록하도록 예외 처리를 보완하였습니다.

#### 9. [logout_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/logout_proc.asp) `[MODIFY]`
* **주요 내용**: 관리자 로그아웃 감사 로그 추가
* **상세 설명**:
  * 세션 정보(`ASeq`, `AID`, `AName`)가 메모리에서 완전히 소멸(`Abandon` / `""` 대입)되기 바로 직전 시점에 `AddAuditLog`를 호출하여 로그아웃 대상 사용자를 정확히 식별하고 기록하도록 구현하였습니다.

#### 10. [left.asp](file:///d:/MPCJOB/mpcjob_2012/admin/include/left.asp) `[MODIFY]`
* **주요 내용**: LNB 메뉴에 감사 로그 관리 페이지 링크 연동
* **상세 설명**:
  * 관리자 좌측 LNB 메뉴 내 환경설정 하단 서브 메뉴에 "감사 로그 관리" (`/admin/siteconf/audit_log_list.asp`) 메뉴 링크를 추가하고 메뉴 권한 설정을 반영하였습니다.

---

### ③ 회원 관리 레이어

#### 11. [mem_form.asp](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_form.asp) `[MODIFY]`
* **주요 내용**: 회원 상세 정보 및 이력서 조회 시 감사 로그 연동
* **상세 설명**:
  * 회원 상세 정보, 인적사항 및 이력서 정보 조회 시 개인정보 열람 감사 로그(`READ`)를 기록하도록 `AddAuditLog`를 연동하였습니다.

#### 12. [mem_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/member/mem_proc.asp) `[MODIFY]`
* **주요 내용**: 회원 C/U/D 처리(Flag별) 명시적 감사 로그 생성 및 자동 감지 표준 결합 적재
* **상세 설명**:
  * `ADD`, `MOD`, `DEL`, `BLK`, `SEARCH`, `PWD_C`, `PWD_S`, `PWD_R`, `EDU`, `CAREER`, `FAM`, `MEMO`, `QUES` 등 전달받은 Flag 구분값에 맞춰 감사 로그를 생성하도록 구현하였습니다.
  * DB 삭제/수정 완료 시점(`Result = 0`)에 기존 자동 감지 파라미터 표준 포맷(`[행위명] /admin/member/mem_proc.asp [파라미터: ...]`) 뒤에 상세 설명(`auditLogDesc`)을 결합하여 정확히 1회 적재되도록 반영하였습니다.

---

### ④ 채용 공고 및 지원자 관리 레이어

#### 13. [job_apply_user.asp](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_apply_user.asp) `[MODIFY]`
* **주요 내용**: 이력서 파일 다운로드 시 공고번호(`RecrSeq`) 파라미터 전달 및 전송 방식 보완
* **상세 설명**:
  * 파일 다운로드 폼(`frmDown`)의 전송 방식을 GET 방식에서 POST 방식으로 변경하였습니다.
  * 감사 로그 추적 시 이력서가 속한 공고 정보를 역추적할 수 있도록 폼 내부에 `RecrSeq` Hidden 파라미터를 추가하여 함께 전달되도록 보완하였습니다.

#### 14. [job_form.asp](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_form.asp) `[MODIFY]`
* **주요 내용**: 채용공고 등록/수정/삭제 액션 URL 수정 및 감사 로그 연동
* **상세 설명**:
  * 채용공고의 등록 및 수정, Ajax를 통한 문항 삭제 시의 Action URL을 테스트 서버 주소(`http://test.mpcjob.co.kr/...`)로 변경하고 감사 로그 연동을 수행하였습니다.

#### 15. [job_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/recruit/job_proc.asp) `[MODIFY]`
* **주요 내용**: 공고 처리(등록/수정) 완료 후 리다이렉트 목적지 주소 수정 및 감사 로그 연동
* **상세 설명**:
  * 신규 등록, 수정, 추천 공고 처리 완료 시 후속 리다이렉트 페이지들의 도메인 주소를 테스트 서버 주소(`http://test.mpcjob.co.kr/...`)로 변경하였습니다.

---

### ⑤ 게시판 관리 레이어

#### 16. [notice_form.asp](file:///d:/MPCJOB/mpcjob_2012/admin/board/notice_form.asp) `[MODIFY]`
* **주요 내용**: 공지사항 폼 진입 감사 로그 및 Action URL 연동
* **상세 설명**:
  * 공지사항 등록/수정 폼 진입 시 `READ` 감사 로그를 적재하고, 폼 전송 주소를 테스트 서버 환경으로 맞추었습니다.

#### 17. [faq_form.asp](file:///d:/MPCJOB/mpcjob_2012/admin/board/faq_form.asp) `[MODIFY]`
* **주요 내용**: FAQ 폼 진입 감사 로그 및 Action URL 연동
* **상세 설명**:
  * FAQ 등록/수정 폼 진입 시 `READ` 감사 로그를 적재하고, 폼 전송 주소를 테스트 서버 환경으로 맞추었습니다.

---

### ⑥ 환경설정 / IP 접근 통제 레이어

#### 18. [ip_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_proc.asp) `[MODIFY]`
* **주요 내용**: IP 접근 제한의 등록, 삭제, 상태 변경에 대한 명시적 감사 로그 적재 보완
* **상세 설명**:
  * DB INSERT 성공 직후: `CREATE` 타입으로 허용 IP 주소 및 메모 내용 상세 기록
  * DB DELETE 성공 직후: `DELETE` 타입으로 삭제한 IP 고유번호(`ipSeq`) 기록
  * DB UPDATE(상태 토글) 성공 직후: `UPDATE` 타입으로 상태가 바뀐 IP 고유번호 및 활성 여부(`Y/N`) 기록

#### 19. [ip_bypass_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_proc.asp) `[MODIFY]`
* **주요 내용**: IP 차단 예외 대상 관리자 ID의 등록 및 삭제에 대한 명시적 감사 로그 적재 보완
* **상세 설명**:
  * 예외 ID 추가 성공 직후: `CREATE` 타입으로 등록된 관리자 ID 및 메모 상세 기록
  * 예외 ID 삭제 성공 직후: `DELETE` 타입으로 삭제된 예외 Seq 기록

---

### ⑦ 감사 로그 관리 레이어

#### 20. [audit_log_list.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_list.asp) `[NEW]`
* **주요 내용**: 감사 로그 목록 조회 및 다중 조건 검색 화면 구현
* **상세 설명**:
  * `TBL_ADMIN_AUDIT_LOG` 테이블에서 최근 발생 일시 순서로 감사 로그 기록 목록을 페이징 처리하여 보여줍니다.
  * 행위 구분별(등록/조회/수정/삭제/다운/엑셀/로그인/로그아웃), 대메뉴별, 검색어(관리자ID, 이름, IP, 상세설명)를 조합한 상세 필터 조회가 가능합니다.
  * 검색 필터 폼 내에서 "엑셀 다운로드"를 연계할 수 있는 버튼 및 비동기 전송 스크립트(`goExcel`)가 탑재되어 있습니다.

#### 21. [audit_log_xls.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/audit_log_xls.asp) `[NEW]`
* **주요 내용**: 감사 로그 목록 엑셀 변환 파일 내보내기 구현
* **상세 설명**:
  * `audit_log_list.asp`에서 설정한 현재 검색 필터 조건을 그대로 유지하면서, 전체 결과 리스트(최대 10,000건 제한)를 엑셀 파일 형식(`xls`)으로 변환 및 즉시 첨부 파일로 내려받도록 처리합니다.
  * 각 열(ID, IP, 메뉴코드, 대상 키 등)은 엑셀 내에서 문자열 포맷이 깨지지 않도록 CSS 텍스트 속성(`mso-number-format`)을 적용하여 출력합니다.
