<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
    ' 1. 권한 체크 (로그인한 일반 관리자 세션이 존재하는지 확인)
    If Session("ASeq") & "" = "" Then
        Response.Write "<script>alert('로그인 후 이용해 주십시오.'); location.href='/admin/login_otp.asp';</script>"
        Response.End
    End If

    Dim listRecord, listPage, Page, RecordCount, PageCount
    Dim startIdx, endIdx
    Dim myIp, listRs, sqlList, arrData, arrDataNum, kData, i, vNum
    myIp = Request.ServerVariables("REMOTE_ADDR")

    listPage   = 10
    listRecord = 10  ' 목록 개수
    Page       = FN_Req("Page", "1")

    If IsNumeric(Page) Then
        Page = CInt(Page)
    Else
        Page = 1
    End If
    If Page < 1 Then Page = 1

    ' DB에서 전체 허용 IP 목록 조회
    On Error Resume Next
    sqlList = "SELECT IpSeq, AllowIp, IpMemo, IsUse, RegDate FROM TBL_IP_ALLOW ORDER BY IpSeq DESC"
    Set listRs = objDbCon.Execute(sqlList)
    If Err.Number <> 0 Then
        Dim errDesc
        errDesc = Err.Description & ""
        errDesc = Replace(errDesc, "'", " ")
        errDesc = Replace(errDesc, """", " ")
        errDesc = Replace(errDesc, vbCrLf, " ")
        Response.Write "<script>alert('DB 조회 오류가 발생했습니다. TBL_IP_ALLOW 테이블이 생성되었는지 확인하십시오.\nError: " & errDesc & "'); history.back();</script>"
        Response.End
    End If
    On Error GoTo 0

    ' GetRows 방식으로 데이터 바인딩 및 페이징 처리
    kData       = False
    arrDataNum  = -1
    RecordCount = 0
    PageCount   = 0

    If Not (listRs Is Nothing) Then
        On Error Resume Next
        If listRs.State = 1 Then
            If Not (listRs.Eof Or listRs.Bof) Then
                arrData     = listRs.GetRows(,,Array("IpSeq","AllowIp","IpMemo","IsUse","RegDate"))
                RecordCount = UBound(arrData, 2) + 1
                PageCount   = FN_PageCount(RecordCount, listRecord)

                If Page > PageCount And PageCount > 0 Then Page = PageCount

                startIdx = (Page - 1) * listRecord
                endIdx   = startIdx + listRecord - 1
                If endIdx > RecordCount - 1 Then endIdx = RecordCount - 1

                arrDataNum  = endIdx
                kData       = True
            End If
            listRs.Close
        End If
        Set listRs = Nothing
        On Error GoTo 0
    End If
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <title>IP 접근 제한 설정 - HC 관리자</title>
    <!--#include virtual="/admin/include/head.asp"-->
    <script type="text/javascript">
        function copyMyIp(ip) {
            $("#allow_ip").val(ip);
        }
        function validateForm() {
            var ip = $.trim($("#allow_ip").val());
            if (ip === "") {
                alert("허용할 IP 주소를 입력해 주십시오.");
                $("#allow_ip").focus();
                return false;
            }
            return true;
        }
        function deleteIp(seq, ip) {
            if (confirm("등록된 IP [" + ip + "] 주소를 목록에서 삭제하시겠습니까?")) {
                location.href = "ip_proc.asp?action=delete&seq=" + seq;
            }
        }
        function toggleUse(seq, currentUse) {
            var nextUse = currentUse === 'Y' ? 'N' : 'Y';
            var msg = nextUse === 'Y' ? '해당 IP 제한을 활성화하시겠습니까?' : '해당 IP 제한을 일시 중지하시겠습니까?';
            if (confirm(msg)) {
                location.href = "ip_proc.asp?action=toggle&seq=" + seq + "&use=" + nextUse;
            }
        }
        function openBypassPop() {
            window.open("ip_bypass_pop.asp", "popBypass", "width=620,height=520,scrollbars=yes,resizable=yes");
        }
    </script>
</head>

<body>
    <div id="wrap">
        <!--#include virtual="/admin/include/top.asp"-->
        
        <div class="contents">
            <!--#include virtual="/admin/include/left.asp"-->
            <div class="right">
                <div class="location"><%=GNB_STR_LOC%></div>
                <div id="contArea">
                    <h2>IP 접근 제한 설정</h2>
                    
                    <!-- IP 추가 및 검색 Area -->
                    <div class="btnArea">
                        <form name="frmIp" id="frmIp" method="post" action="ip_proc.asp" onsubmit="return validateForm();">
                            <input type="hidden" name="action" value="insert" />
                            
                            <div class="fl_l" style="line-height:30px; font-size:12px; color:#475569;">
                                현재 접속 IP: <strong style="color:#1d4ed8; font-size:13px;"><%=myIp%></strong>
                                <span class="btns btn_g" style="margin-left:10px;">
                                    <a href="javascript:void(0);" onclick="copyMyIp('<%=myIp%>')">내 IP 자동입력</a>
                                </span>
                                <span class="btns btn_g" style="margin-left:5px;">
                                    <a href="javascript:void(0);" onclick="openBypassPop();">🔓 예외 ID 관리</a>
                                </span>
                            </div>
                            
                            <div class="fl_r" style="font-size:12px;">
                                ■ 허용 IP 주소: 
                                <input type="text" id="allow_ip" name="allow_ip" class="w130" placeholder="예: 210.112.73.45" />
                                
                                ■ 식별 메모: 
                                <input type="text" id="ip_memo" name="ip_memo" class="w200" placeholder="용도 구분용 메모" />
                                
                                <span class="btns btn_b">
                                    <button type="submit">허용 IP 등록</button>
                                </span>
                            </div>
                        </form>
                    </div>
                    
                    <!-- IP 목록 리스트 Area -->
                    <div class="tblArea">
                        <table>
                            <colgroup>
                                <col class="w50" />
                                <col class="w180" />
                                <col />
                                <col class="w110" />
                                <col class="w160" />
                                <col class="w90" />
                            </colgroup>
                            <thead>
                                <tr class="design">
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                </tr>
                                <tr>
                                    <th>번호</th>
                                    <th>허용 IP 주소</th>
                                    <th>식별용 메모</th>
                                    <th>사용 상태</th>
                                    <th>등록일시</th>
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% If Not kData Then %>
                                <tr>
                                    <td colspan="6" style="padding: 40px 0; color: #94a3b8;">
                                        등록된 허용 IP 주소가 없습니다.<br/>
                                        (IP 제한이 작동하지 않아 현재 모든 대역에서 접속이 가능합니다.)
                                    </td>
                                </tr>
                                <% Else %>
                                <%      
                                        vNum = RecordCount - ((Page - 1) * listRecord)
                                        For i = startIdx To endIdx 
                                %>
                                <tr>
                                    <td><%=vNum%></td>
                                    <td style="font-weight: bold; font-family: monospace; font-size: 14px;"><%=arrData(1,i)%></td>
                                    <td class="alignL"><%=arrData(2,i)%></td>
                                    <td>
                                        <a href="javascript:void(0);" onclick="toggleUse('<%=arrData(0,i)%>', '<%=arrData(3,i)%>')">
                                        <% If arrData(3,i) = "Y" Then %>
                                            <strong style="color:#16a34a;">사용 중</strong>
                                        <% Else %>
                                            <span style="color:#dc2626;">일시 중지</span>
                                        <% End If %>
                                        </a>
                                    </td>
                                    <td><%=arrData(4,i)%></td>
                                    <td>
                                        <span class="btns btn_g">
                                            <a href="javascript:void(0);" onclick="deleteIp('<%=arrData(0,i)%>', '<%=arrData(1,i)%>')">삭제</a>
                                        </span>
                                    </td>
                                </tr>
                                <%          
                                            vNum = vNum - 1
                                        Next 
                                %>
                                <% End If %>
                            </tbody>
                        </table>

                        <form name="frmPage" id="frmPage" method="post" action="ip_list.asp">
                            <input type="hidden" id="Page" name="Page" value="<%=Page%>" />
                        </form>
                        
                        <div class="paging">
                            <% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!--#include virtual="/admin/include/footer.asp"-->
    </div>
</body>
</html>
