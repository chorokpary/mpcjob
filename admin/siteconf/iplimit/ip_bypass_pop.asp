<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
    ' 권한 체크 (로그인 세션 확인)
    If Session("ASeq") & "" = "" Then
        Response.Write "<script>alert('로그인 후 이용해 주십시오.'); window.close();</script>"
        Response.End
    End If

    Dim bypassRs, sqlList, arrData, arrDataNum, kData, i, vNum, RecordCount
    
    ' DB에서 예외 허용 ID 목록 조회
    On Error Resume Next
    sqlList = "SELECT BypassSeq, AdminId, BypassMemo, RegDate FROM TBL_IP_BYPASS ORDER BY BypassSeq DESC"
    Set bypassRs = objDbCon.Execute(sqlList)
    If Err.Number <> 0 Then
        Dim errDesc
        errDesc = Err.Description & ""
        errDesc = Replace(errDesc, "'", " ")
        errDesc = Replace(errDesc, """", " ")
        errDesc = Replace(errDesc, vbCrLf, " ")
        Response.Write "<script>alert('DB 조회 오류가 발생했습니다. TBL_IP_BYPASS 테이블이 생성되었는지 확인하십시오.\nError: " & errDesc & "');</script>"
    End If
    On Error GoTo 0

    kData       = False
    arrDataNum  = -1
    RecordCount = 0

    If Not (bypassRs Is Nothing) Then
        On Error Resume Next
        If bypassRs.State = 1 Then
            If Not (bypassRs.Eof Or bypassRs.Bof) Then
                arrData     = bypassRs.GetRows(,,Array("BypassSeq","AdminId","BypassMemo","RegDate"))
                arrDataNum  = UBound(arrData, 2)
                kData       = True
                RecordCount = arrDataNum + 1
            End If
            bypassRs.Close
        End If
        Set bypassRs = Nothing
        On Error GoTo 0
    End If
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8" />
    <title>예외 관리자 ID 관리</title>
    <link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
    <script type="text/javascript">
        function validateForm() {
            var adminId = $.trim($("#admin_id").val());
            if (adminId === "") {
                alert("예외 처리할 관리자 ID를 입력해 주십시오.");
                $("#admin_id").focus();
                return false;
            }
            return true;
        }
        function deleteBypass(seq, adminId) {
            if (confirm("등록된 예외 ID [" + adminId + "]를 목록에서 삭제하시겠습니까?")) {
                location.href = "ip_bypass_proc.asp?action=delete&seq=" + seq;
            }
        }
    </script>
    <style type="text/css">
        body { padding: 15px; background: #f8fafc; font-family: dotum, sans-serif; font-size: 12px; }
        .pop_title { font-size: 16px; font-weight: bold; color: #1e293b; padding-bottom: 10px; border-bottom: 2px solid #0284c7; margin-bottom: 15px; }
        .reg_box { background: #ffffff; padding: 12px; border: 1px solid #cbd5e1; border-radius: 4px; margin-bottom: 15px; }
        .reg_box input { height: 24px; line-height: 24px; padding: 0 5px; border: 1px solid #cbd5e1; border-radius: 3px; }
        table { width: 100%; border-collapse: collapse; background: #ffffff; }
        th { background: #f1f5f9; padding: 8px; border: 1px solid #cbd5e1; color: #334155; font-size: 12px; }
        td { padding: 8px; border: 1px solid #e2e8f0; text-align: center; color: #475569; }
        .btn_sub { background: #0284c7; color: #fff; border: none; padding: 4px 10px; cursor: pointer; border-radius: 3px; font-weight: bold; }
        .btn_del { background: #ef4444; color: #fff; border: none; padding: 2px 8px; cursor: pointer; border-radius: 3px; font-size: 11px; }
    </style>
</head>
<body>
    <div class="pop_title">IP 접근 제한 예외 관리자 ID 설정</div>
    
    <!-- ID 등록 폼 -->
    <div class="reg_box">
        <form name="frmBypass" id="frmBypass" method="post" action="ip_bypass_proc.asp" onsubmit="return validateForm();">
            <input type="hidden" name="action" value="insert" />
            <strong>■ 예외 ID:</strong> 
            <input type="text" id="admin_id" name="admin_id" style="width:110px;" placeholder="관리자 ID" />
            
            <strong style="margin-left:8px;">■ 메모:</strong> 
            <input type="text" id="bypass_memo" name="bypass_memo" style="width:160px;" placeholder="구분용 메모" />
            
            <button type="submit" class="btn_sub" style="margin-left:5px;">등록</button>
        </form>
    </div>

    <!-- 예외 ID 목록 테이블 -->
    <table>
        <thead>
            <tr>
                <th style="width:40px;">번호</th>
                <th style="width:120px;">예외 관리자 ID</th>
                <th>메모</th>
                <th style="width:130px;">등록일시</th>
                <th style="width:50px;">관리</th>
            </tr>
        </thead>
        <tbody>
            <% If Not kData Then %>
            <tr>
                <td colspan="5" style="padding:25px 0; color:#94a3b8;">
                    등록된 예외 관리자 ID가 없습니다.
                </td>
            </tr>
            <% Else %>
            <%      
                    vNum = RecordCount
                    For i = 0 To arrDataNum 
            %>
            <tr>
                <td><%=vNum%></td>
                <td style="font-weight:bold; color:#0369a1;"><%=arrData(1,i)%></td>
                <td style="text-align:left;"><%=arrData(2,i)%></td>
                <td><%=arrData(3,i)%></td>
                <td>
                    <button type="button" class="btn_del" onclick="deleteBypass('<%=arrData(0,i)%>', '<%=arrData(1,i)%>')">삭제</button>
                </td>
            </tr>
            <%          
                        vNum = vNum - 1
                    Next 
            %>
            <% End If %>
        </tbody>
    </table>

    <div style="text-align:center; margin-top:15px;">
        <button type="button" onclick="window.close();" style="padding:6px 20px; background:#64748b; color:#fff; border:none; border-radius:3px; cursor:pointer;">창닫기</button>
    </div>
</body>
</html>
