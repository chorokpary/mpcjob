<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
    ' 1. 권한 체크 (로그인한 일반 관리자 세션이 존재하는지 확인)
    If Session("ASeq") & "" = "" Then
        Response.Write "<script>alert('로그인 후 이용해 주십시오.'); location.href='/admin/login_otp.asp';</script>"
        Response.End
    End If

    Response.Buffer = TRUE
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition","attachment; filename=Audit_Log_" & Replace(Date(),"-","") & ".xls"
%>
<%
    Dim listRs, sqlList, arrData, arrDataNum, kData, i
    Dim sqlWhere, sLogType, sParentMenu, sSearch, sBDate, sEDate

    ' 검색 인자 받기
    sLogType    = Trim(FN_Req("sLogType", ""))
    sParentMenu = Trim(FN_Req("sParentMenu", ""))
    sSearch     = Trim(FN_Req("sSearch", ""))
    sBDate      = Trim(FN_Req("sBDate", ""))
    sEDate      = Trim(FN_Req("sEDate", ""))

    ' Where 조건 구성
    sqlWhere = " WHERE 1=1 "
    If sBDate <> "" Then
        sqlWhere = sqlWhere & " AND RegDate >= '" & Replace(sBDate, "'", "''") & " 00:00:00' "
    End If
    If sEDate <> "" Then
        sqlWhere = sqlWhere & " AND RegDate <= '" & Replace(sEDate, "'", "''") & " 23:59:59' "
    End If
    If sLogType <> "" Then
        sqlWhere = sqlWhere & " AND LogType = '" & Replace(sLogType, "'", "''") & "' "
    End If
    If sParentMenu <> "" Then
        sqlWhere = sqlWhere & " AND ParentMenuName = '" & Replace(sParentMenu, "'", "''") & "' "
    End If
    If sSearch <> "" Then
        Dim safeSearch : safeSearch = Replace(sSearch, "'", "''")
        sqlWhere = sqlWhere & " AND (AdminID LIKE '%" & safeSearch & "%' OR AdminName LIKE '%" & safeSearch & "%' OR AdminIP LIKE '%" & safeSearch & "%' OR LogDesc LIKE '%" & safeSearch & "%') "
    End If

    ' DB 조회 (최대 10000개 제한)
    On Error Resume Next
    sqlList = "SELECT TOP 10000 LogSeq, AdminSeq, AdminID, AdminName, AdminIP, MenuCode, ParentMenuName, MenuName, SubMenuName, LogType, TargetKey, LogDesc, RegDate " & _
              "FROM TBL_ADMIN_AUDIT_LOG " & sqlWhere & " ORDER BY LogSeq DESC"
    Set listRs = objDbCon.Execute(sqlList)
    If Err.Number <> 0 Then
        Response.Write "DB Error: " & Err.Description
        Response.End
    End If
    On Error GoTo 0

    kData       = False
    arrDataNum  = -1

    If Not (listRs Is Nothing) Then
        If listRs.State = 1 Then
            If Not (listRs.Eof Or listRs.Bof) Then
                arrData     = listRs.GetRows(,,Array("LogSeq", "AdminSeq", "AdminID", "AdminName", "AdminIP", "MenuCode", "ParentMenuName", "MenuName", "SubMenuName", "LogType", "TargetKey", "LogDesc", "RegDate"))
                arrDataNum  = UBound(arrData, 2)
                kData       = True
            End If
            listRs.Close
        End If
        Set listRs = Nothing
    End If
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
    <style>
        .tbl_xls { border-collapse: collapse; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; }
        .tbl_xls th { background-color: #f1f5f9; border: 1px solid #cbd5e1; height: 30px; font-weight: bold; }
        .tbl_xls td { border: 1px solid #cbd5e1; height: 25px; text-align: center; }
        .tbl_xls td.left { text-align: left; }
    </style>
</head>
<body>
    <table class="tbl_xls">
        <thead>
            <tr>
                <th>번호</th>
                <th>발생일시</th>
                <th>관리자 ID</th>
                <th>관리자명</th>
                <th>관리자 IP</th>
                <th>대메뉴</th>
                <th>소메뉴</th>
                <th>구분</th>
                <th>고유키값</th>
                <th>작업 상세 내용</th>
            </tr>
        </thead>
        <tbody>
            <% If Not kData Then %>
            <tr>
                <td colspan="10" style="height:50px; text-align:center;">조회된 감사 로그 내역이 없습니다.</td>
            </tr>
            <% Else %>
            <%
                For i = 0 To arrDataNum
                    Dim logTypeStr
                    Select Case arrData(9,i)
                        Case "CREATE"     : logTypeStr = "등록"
                        Case "READ"       : logTypeStr = "조회"
                        Case "UPDATE"     : logTypeStr = "수정"
                        Case "DELETE"     : logTypeStr = "삭제"
                        Case "FILEDOWN"   : logTypeStr = "다운"
                        Case "EXCEL"      : logTypeStr = "엑셀"
                        Case "LOGIN"      : logTypeStr = "로그인"
                        Case "LOGIN_FAIL" : logTypeStr = "로그인 실패"
                        Case "LOGOUT"     : logTypeStr = "로그아웃"
                        Case Else         : logTypeStr = arrData(9,i)
                    End Select
            %>
            <tr>
                <td><%=arrDataNum - i + 1%></td>
                <td><%=arrData(12,i)%></td>
                <td style="mso-number-format:'\@';"><%=arrData(2,i)%></td>
                <td><%=arrData(3,i)%></td>
                <td style="mso-number-format:'\@';"><%=arrData(4,i)%></td>
                <td><%=arrData(6,i)%></td>
                <td><%=arrData(7,i)%></td>
                <td><%=logTypeStr%></td>
                <td style="mso-number-format:'\@';"><%=arrData(10,i)%></td>
                <td class="left"><%=Trim(Replace(arrData(11,i), "자동감지/", ""))%></td>
            </tr>
            <%
                Next
            End If
            %>
        </tbody>
    </table>
</body>
</html>
