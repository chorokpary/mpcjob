<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 1. 권한 체크 (로그인한 일반 관리자 세션이 존재하는지 확인)
    If Session("ASeq") = "" Or IsNull(Session("ASeq")) Then
        Response.Write "<script>alert('로그인 후 이용해 주십시오.'); location.href='/admin/login_otp.asp';</script>"
        Response.End
    End If

    Dim action, ipSeq, allowIp, ipMemo, isUse, sqlQuery, rowsAffected
    action = Trim(Request("action"))

    Select Case action
        ' ----------------------------------------------------
        ' 1. 신규 허용 IP 등록
        ' ----------------------------------------------------
        Case "insert"
            allowIp = Trim(Request.Form("allow_ip"))
            ipMemo  = Trim(Request.Form("ip_memo"))
            
            ' 입력값 유효성 체크
            If allowIp = "" Then
                Response.Write "<script>alert('IP 주소를 입력해 주십시오.'); history.back();</script>"
                Response.End
            End If

            ' SQL Injection 방지 문자열 필터링
            allowIp = Replace(allowIp, "'", "''")
            ipMemo  = Replace(ipMemo, "'", "''")

            On Error Resume Next
            sqlQuery = "INSERT INTO TBL_IP_ALLOW (AllowIp, IpMemo, IsUse, RegDate) VALUES (" & _
                       "'" & allowIp & "', " & _
                       "'" & ipMemo & "', " & _
                       "'Y', GETDATE())"
            objDbCon.Execute sqlQuery, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('IP 등록 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "ip_list.asp"
            Response.End

        ' ----------------------------------------------------
        ' 2. 등록된 허용 IP 삭제
        ' ----------------------------------------------------
        Case "delete"
            ipSeq = Trim(Request("seq"))
            
            If ipSeq = "" Or Not IsNumeric(ipSeq) Then
                Response.Write "<script>alert('올바르지 않은 요청입니다.'); history.back();</script>"
                Response.End
            End If

            On Error Resume Next
            sqlQuery = "DELETE FROM TBL_IP_ALLOW WHERE IpSeq = " & ipSeq
            objDbCon.Execute sqlQuery, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('IP 삭제 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "ip_list.asp"
            Response.End

        ' ----------------------------------------------------
        ' 3. IP 사용 상태 토글 (활성/비활성)
        ' ----------------------------------------------------
        Case "toggle"
            ipSeq = Trim(Request("seq"))
            isUse = Trim(Request("use"))

            If ipSeq = "" Or Not IsNumeric(ipSeq) Or (isUse <> "Y" And isUse <> "N") Then
                Response.Write "<script>alert('올바르지 않은 요청입니다.'); history.back();</script>"
                Response.End
            End If

            On Error Resume Next
            sqlQuery = "UPDATE TBL_IP_ALLOW SET IsUse = '" & isUse & "' WHERE IpSeq = " & ipSeq
            objDbCon.Execute sqlQuery, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('상태 변경 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "ip_list.asp"
            Response.End

        ' ----------------------------------------------------
        ' 비정상 접근 분기
        ' ----------------------------------------------------
        Case Else
            Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='ip_list.asp';</script>"
            Response.End
    End Select
%>
