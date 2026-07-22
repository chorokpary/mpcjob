<!-- #include virtual="/common/CommonConfig.asp" -->
<%
    ' 권한 체크
    If Session("ASeq") & "" = "" Then
        Response.Write "<script>alert('로그인 후 이용해 주십시오.'); window.close();</script>"
        Response.End
    End If

    Dim action, bypassSeq, adminId, bypassMemo, sqlQuery, rowsAffected
    action = Trim(Request("action"))

    Select Case action
        ' ----------------------------------------------------
        ' 1. 예외 관리자 ID 등록
        ' ----------------------------------------------------
        Case "insert"
            adminId    = Trim(Request.Form("admin_id"))
            bypassMemo = Trim(Request.Form("bypass_memo"))
            
            If adminId = "" Then
                Response.Write "<script>alert('예외 처리할 관리자 ID를 입력해 주십시오.'); history.back();</script>"
                Response.End
            End If

            adminId    = Replace(adminId, "'", "''")
            bypassMemo = Replace(bypassMemo, "'", "''")

            On Error Resume Next
            sqlQuery = "INSERT INTO TBL_IP_BYPASS (AdminId, BypassMemo, RegDate) VALUES (" & _
                       "'" & adminId & "', " & _
                       "'" & bypassMemo & "', " & _
                       "GETDATE())"
            objDbCon.Execute sqlQuery, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('예외 ID 등록 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "ip_bypass_pop.asp"
            Response.End

        ' ----------------------------------------------------
        ' 2. 예외 관리자 ID 삭제
        ' ----------------------------------------------------
        Case "delete"
            bypassSeq = Trim(Request("seq"))
            
            If bypassSeq = "" Or Not IsNumeric(bypassSeq) Then
                Response.Write "<script>alert('올바르지 않은 요청입니다.'); history.back();</script>"
                Response.End
            End If

            On Error Resume Next
            sqlQuery = "DELETE FROM TBL_IP_BYPASS WHERE BypassSeq = " & bypassSeq
            objDbCon.Execute sqlQuery, rowsAffected
            
            If Err.Number <> 0 Then
                Response.Write "<script>alert('삭제 도중 오류가 발생했습니다.\nError: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "ip_bypass_pop.asp"
            Response.End

        Case Else
            Response.Write "<script>alert('비정상적인 접근입니다.'); location.href='ip_bypass_pop.asp';</script>"
            Response.End
    End Select
%>
