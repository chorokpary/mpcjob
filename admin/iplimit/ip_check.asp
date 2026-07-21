<%
    ' ====================================================================
    ' * Description : ip_check.asp / 관리자 접속 IP 접근 제한 체크 모듈
    ' ====================================================================
    
    Dim userIp, ipCheckRs, activeIpCount, isAllowedIp, checkSql
    userIp = Request.ServerVariables("REMOTE_ADDR")
    
    Dim bypassIds, curAid, isBypassId, xId, bypassStr
    
    ' DB 테이블(TBL_IP_BYPASS)에서 예외 ID 목록 조회하여 bypassIds 배열 생성
    bypassStr = "testadmin" ' 하드코딩 기본 우회 ID
    On Error Resume Next
    Set ipCheckRs = objDbCon.Execute("SELECT AdminId FROM TBL_IP_BYPASS")
    If Err.Number = 0 Then
        Do While Not ipCheckRs.Eof
            If Trim(ipCheckRs("AdminId") & "") <> "" Then
                If bypassStr <> "" Then bypassStr = bypassStr & ","
                bypassStr = bypassStr & LCase(Trim(ipCheckRs("AdminId") & ""))
            End If
            ipCheckRs.MoveNext
        Loop
        ipCheckRs.Close
        Set ipCheckRs = Nothing
    End If
    On Error GoTo 0

    bypassIds = Split(bypassStr, ",")

    ' 현재 접속/로그인 시도 계정 ID 확인 (Session 및 로그인 폼 전달값 감지)
    curAid = LCase(Session("AID") & "")
    If curAid = "" Then curAid = LCase(Session("TempAID") & "")
    If curAid = "" Then curAid = LCase(Trim(Request("txtID") & ""))
    
    isBypassId = False
    If curAid <> "" Then
        For Each xId In bypassIds
            If curAid = LCase(Trim(xId)) Then
                isBypassId = True
                Exit For
            End If
        Next
    End If
    
    ' 예외 허용된 관리자 계정(ID)은 무조건 통과
    If isBypassId Then
        ' 통과
    Else
        activeIpCount = 0
        isAllowedIp = False
        
        ' 1. 사용 중인(활성화된) 허용 IP 목록이 존재하는지 체크
        On Error Resume Next
        Set ipCheckRs = objDbCon.Execute("SELECT COUNT(*) FROM TBL_IP_ALLOW WHERE IsUse = 'Y'")
        If Err.Number = 0 Then
            If Not ipCheckRs.Eof Then
                activeIpCount = CInt(ipCheckRs(0))
            End If
            ipCheckRs.Close
            Set ipCheckRs = Nothing
        End If
        On Error GoTo 0
        
        ' 활성화된 허용 IP가 1개 이상 존재할 때만 차단 작동 (Lockout 방지)
        If activeIpCount > 0 Then
            On Error Resume Next
            ' 현재 접속자의 IP가 허용 IP 테이블에 등록되어 있고 활성화 상태인지 확인 (와일드카드 * 지원)
            checkSql = "SELECT COUNT(*) FROM TBL_IP_ALLOW WHERE IsUse = 'Y' AND ('" & Replace(userIp, "'", "''") & "' LIKE REPLACE(AllowIp, '*', '%'))"
            Set ipCheckRs = objDbCon.Execute(checkSql)
            If Err.Number = 0 Then
                If Not ipCheckRs.Eof Then
                    If CInt(ipCheckRs(0)) > 0 Then
                        isAllowedIp = True
                    End If
                End If
                ipCheckRs.Close
                Set ipCheckRs = Nothing
            End If
            On Error GoTo 0
            
            ' 허용되지 않은 IP인 경우 경고 후 튕겨냄
            If Not isAllowedIp Then
                Response.ContentType = "text/html"
                Response.CharSet = "utf-8"
                Response.Write "<script type=""text/javascript"">"
                Response.Write "alert('허용되지 않은 IP 주소(" & userIp & ")에서 접속을 시도했습니다.\n관리자에게 문의하십시오.');"
                Response.Write "location.href = '/admin/login.asp';"
                Response.Write "</script>"
                Response.End
            End If
        End If
    End If
%>
