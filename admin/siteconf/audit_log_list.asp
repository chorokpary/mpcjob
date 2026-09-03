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
    Dim listRs, sqlList, arrData, arrDataNum, kData, i, vNum
    Dim sqlWhere, sLogType, sParentMenu, sSearch, sBDate, sEDate

    listPage   = 10
    listRecord = FN_Req("ListSize", "20")
    If IsNumeric(listRecord) Then
        listRecord = CInt(listRecord)
    Else
        listRecord = 20
    End If
    If listRecord < 5 Then listRecord = 20

    Page       = FN_Req("Page", "1")

    If IsNumeric(Page) Then
        Page = CInt(Page)
    Else
        Page = 1
    End If
    If Page < 1 Then Page = 1

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

    ' DB 조회
    On Error Resume Next
    sqlList = "SELECT LogSeq, AdminSeq, AdminID, AdminName, AdminIP, MenuCode, ParentMenuName, MenuName, SubMenuName, LogType, TargetKey, LogDesc, RegDate " & _
              "FROM TBL_ADMIN_AUDIT_LOG " & sqlWhere & " ORDER BY LogSeq DESC"
    Set listRs = objDbCon.Execute(sqlList)
    If Err.Number <> 0 Then
        Dim errDesc : errDesc = Err.Description & ""
        errDesc = Replace(errDesc, "'", " ")
        errDesc = Replace(errDesc, """", " ")
        errDesc = Replace(errDesc, vbCrLf, " ")
        Response.Write "<script>alert('DB 조회 오류가 발생했습니다. TBL_ADMIN_AUDIT_LOG 테이블이 존재하지 않거나 구성이 잘못되었는지 확인하십시오.\nError: " & errDesc & "'); history.back();</script>"
        Response.End
    End If
    On Error GoTo 0

    ' 데이터 바인딩 및 페이징
    kData       = False
    arrDataNum  = -1
    RecordCount = 0
    PageCount   = 0

    If Not (listRs Is Nothing) Then
        On Error Resume Next
        If listRs.State = 1 Then
            If Not (listRs.Eof Or listRs.Bof) Then
                arrData     = listRs.GetRows(,,Array("LogSeq", "AdminSeq", "AdminID", "AdminName", "AdminIP", "MenuCode", "ParentMenuName", "MenuName", "SubMenuName", "LogType", "TargetKey", "LogDesc", "RegDate"))
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
    <title>감사 로그 관리 - HC 관리자</title>
    <!--#include virtual="/admin/include/head.asp"-->
    <link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
    <style type="text/css">
        .desc-tooltip-container { position: relative; display: inline-block; }
        .desc-tooltip {
            display: none;
            position: absolute;
            top: 25px;
            left: 50%;
            transform: translateX(-50%);
            width: 320px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.1);
            border-radius: 6px;
            padding: 12px;
            text-align: left;
            z-index: 999;
            color: #334155;
            line-height: 1.5;
            font-size: 11px;
            word-break: break-all;
        }
        .desc-tooltip .tooltip-close {
            position: absolute;
            top: 8px;
            right: 8px;
            font-weight: bold;
            cursor: pointer;
            color: #94a3b8;
            font-size: 16px;
            line-height: 1;
        }
        .desc-tooltip .tooltip-close:hover { color: #64748b; }
    </style>
    <script type="text/javascript">
        function goPage(page) {
            $("#Page").val(page);
            $("#frmSearch").submit();
        }
        function goExcel() {
            var frm = document.frmSearch;
            var originalAction = frm.action;
            frm.action = "audit_log_xls.asp";
            frm.submit();
            frm.action = originalAction;
        }
        $(function() {
            $("#ListSize").change(function(){ goPage(1); });
            $("#sBDate").datepicker({
                date: $("#sBDate").val()
                , current: $("#sBDate").val()
            });
            $("#sEDate").datepicker({
                date: $("#sEDate").val()
                , current: $("#sEDate").val()
            });

            // 상세보기 클릭 시 툴팁 토글
            $(document).on("click", ".btnShowDesc", function(e) {
                e.stopPropagation();
                var $tooltip = $(this).closest(".desc-tooltip-container").find(".desc-tooltip");
                $(".desc-tooltip").not($tooltip).hide();
                $tooltip.toggle();
            });

            // 툴팁 닫기
            $(document).on("click", ".tooltip-close", function(e) {
                e.stopPropagation();
                $(this).closest(".desc-tooltip").hide();
            });

            // 영역 외 클릭 시 닫기
            $(document).click(function(e) {
                if (!$(e.target).closest(".desc-tooltip-container").length) {
                    $(".desc-tooltip").hide();
                }
            });
        });
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
                    <h2>감사 로그 관리</h2>
                    
                    <form name="frmSearch" id="frmSearch" method="post" action="audit_log_list.asp">
                        <input type="hidden" name="Page" id="Page" value="<%=Page%>" />
                        
                        <div class="searchArea">
                            <!-- 첫 번째 줄: 기간 검색 -->
                            <div class="search_date" style="margin-bottom:10px;">
                                ■ 기간 : 
                                <input type="text" id="sBDate" name="sBDate" value="<%=sBDate%>" class="w80" style="margin-right:5px;" maxlength="10" /> 부터 ~ 
                                <input type="text" id="sEDate" name="sEDate" value="<%=sEDate%>" class="w80" style="margin-right:5px;" maxlength="10" /> 까지
                            </div>
                            
                            <!-- 두 번째 줄: 상세 필터 및 검색 -->
                            <div class="search_date">
                                ■ 구분 : 
                                <select name="sLogType" id="sLogType" title="구분" class="w110" onchange="goPage(1);">
                                    <option value="">전체</option>
                                    <option value="CREATE" <%=Fn_SetDefault(sLogType, "CREATE", "selected", "")%>>등록</option>
                                    <option value="READ" <%=Fn_SetDefault(sLogType, "READ", "selected", "")%>>조회</option>
                                    <option value="UPDATE" <%=Fn_SetDefault(sLogType, "UPDATE", "selected", "")%>>수정</option>
                                    <option value="DELETE" <%=Fn_SetDefault(sLogType, "DELETE", "selected", "")%>>삭제</option>
                                    <option value="FILEDOWN" <%=Fn_SetDefault(sLogType, "FILEDOWN", "selected", "")%>>다운</option>
                                    <option value="EXCEL" <%=Fn_SetDefault(sLogType, "EXCEL", "selected", "")%>>엑셀</option>
                                    <option value="LOGIN" <%=Fn_SetDefault(sLogType, "LOGIN", "selected", "")%>>로그인 성공</option>
                                    <option value="LOGOUT" <%=Fn_SetDefault(sLogType, "LOGOUT", "selected", "")%>>로그아웃</option>
                                </select>
                                
                                / 메뉴 : 
                                <select name="sParentMenu" id="sParentMenu" title="메뉴" class="w110" onchange="goPage(1);">
                                    <option value="">전체</option>
                                    <option value="채용관리" <%=Fn_SetDefault(sParentMenu, "채용관리", "selected", "")%>>채용관리</option>
                                    <option value="회원관리" <%=Fn_SetDefault(sParentMenu, "회원관리", "selected", "")%>>회원관리</option>
                                    <option value="프로젝트관리" <%=Fn_SetDefault(sParentMenu, "프로젝트관리", "selected", "")%>>프로젝트관리</option>
                                    <option value="게시글관리" <%=Fn_SetDefault(sParentMenu, "게시글관리", "selected", "")%>>게시글관리</option>
                                    <option value="통계관리" <%=Fn_SetDefault(sParentMenu, "통계관리", "selected", "")%>>통계관리</option>
                                    <option value="환경설정" <%=Fn_SetDefault(sParentMenu, "환경설정", "selected", "")%>>환경설정</option>
                                    <option value="쪽지함" <%=Fn_SetDefault(sParentMenu, "쪽지함", "selected", "")%>>쪽지함</option>
                                </select>
                                
                                / 통합검색 : 
                                <input type="text" id="sSearch" name="sSearch" value="<%=sSearch%>" title="검색어" class="w190" placeholder="ID/이름/IP/상세내용" />
                                
                                <span class="btns btn_b">
                                    <button type="button" onclick="goPage(1);">검색</button>
                                </span>
                            </div>
                        </div>

                        <!-- 기능 버튼 및 목록 사이즈 제어 Area -->
                        <div class="btnArea">
                            <div class="fl_l">
                                <select name="ListSize" id="ListSize" title="출력개수" class="w80">
                                    <option value="20">출력개수</option>
                                    <option value="10"<%=Fn_SetDefault(listRecord,"10"," selected","")%>>10개씩</option>
                                    <option value="20"<%=Fn_SetDefault(listRecord,"20"," selected","")%>>20개씩</option>
                                    <option value="30"<%=Fn_SetDefault(listRecord,"30"," selected","")%>>30개씩</option>
                                    <option value="50"<%=Fn_SetDefault(listRecord,"50"," selected","")%>>50개씩</option>
                                    <option value="100"<%=Fn_SetDefault(listRecord,"100"," selected","")%>>100개씩</option>
                                </select>
                            </div>
                            <div class="fl_r">
                                <span class="btns btn_b"><a href="javascript:void(0);" onclick="goExcel();">엑셀 다운로드</a></span>
                            </div>
                        </div>
                    </form>
                    
                    <!-- 감사 로그 리스트 Area -->
                    <div class="tblArea">
                        <table>
                            <colgroup>
                                <col class="w50" />
                                <col class="w140" />
                                <col class="w100" />
                                <col class="w80" />
                                <col class="w90" />
                                <col />
                                <col class="w150" />
                                <col class="w100" />
                            </colgroup>
                            <thead>
                                <tr class="design">
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                </tr>
                                <tr>
                                    <th>번호</th>
                                    <th>발생 일시</th>
                                    <th>관리자 ID</th>
                                    <th>관리자명</th>
                                    <th>관리자 IP</th>
                                    <th>메뉴명</th>
                                    <th>구분</th>
                                    <th>작업 상세 내용</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% If Not kData Then %>
                                <tr>
                                    <td colspan="8" style="padding: 40px 0; color: #94a3b8;">
                                        조회된 감사 로그 내역이 없습니다.
                                    </td>
                                </tr>
                                <% Else %>
                                <%      
                                        vNum = RecordCount - ((Page - 1) * listRecord)
                                        For i = startIdx To endIdx 
                                            Dim logTypeStr, logTypeColor
                                            Select Case arrData(9,i)
                                                Case "CREATE"     : logTypeStr = "[등록]" : logTypeColor = "#16a34a"
                                                Case "READ"       : logTypeStr = "[조회]" : logTypeColor = "#2563eb"
                                                Case "UPDATE"     : logTypeStr = "[수정]" : logTypeColor = "#ea580c"
                                                Case "DELETE"     : logTypeStr = "[삭제]" : logTypeColor = "#dc2626"
                                                Case "FILEDOWN"   : logTypeStr = "[다운]" : logTypeColor = "#0891b2"
                                                Case "EXCEL"      : logTypeStr = "[엑셀]" : logTypeColor = "#0d9488"
                                                Case "LOGIN"      : logTypeStr = "[로그인]" : logTypeColor = "#4f46e5"
                                                Case "LOGIN_FAIL" : logTypeStr = "[로그인실패]" : logTypeColor = "#7c3aed"
                                                Case "LOGOUT"     : logTypeStr = "[로그아웃]" : logTypeColor = "#475569"
                                                Case Else         : logTypeStr = "[" & arrData(9,i) & "]" : logTypeColor = "#64748b"
                                            End Select

                                            Dim dispMenuName
                                            If arrData(6,i) <> "" Then
                                                dispMenuName = arrData(6,i) & " > " & arrData(7,i)
                                            Else
                                                dispMenuName = arrData(7,i)
                                            End If
                                %>
                                <tr>
                                    <td><%=vNum%></td>
                                    <td style="font-size: 11px; color: #64748b;"><%=FN_SetDateTimeFormat(arrData(12,i), "YYYY-MM-DD HH:NN:SS")%></td>
                                    <td style="font-family: monospace; font-size: 11px;"><%=arrData(2,i)%></td>
                                    <td><strong><%=arrData(3,i)%></strong></td>
                                    <td style="font-family: monospace; font-size: 11px; color: #64748b;"><%=arrData(4,i)%></td>
                                    <td><%=dispMenuName%></td>
                                    <td>
                                        <strong style="color:<%=logTypeColor%>;"><%=logTypeStr%></strong>
                                    </td>
                                    <td style="line-height: 1.5; font-size:12px; position: relative; text-align: center;">
                                        <div class="desc-tooltip-container">
                                            <span class="btns btn_b" style="display: inline-block;">
                                                <button type="button" class="btnShowDesc" style="padding: 2px 6px; font-size: 11px;">상세 보기</button>
                                            </span>
                                            
                                            <div class="desc-tooltip">
                                                <span class="tooltip-close">&times;</span>
                                                <strong style="display:block; margin-bottom:8px; color:#1e293b; font-size:12px; border-bottom:1px solid #e2e8f0; padding-bottom:5px;">작업 상세 내용</strong>
                                                <div style="max-height: 160px; overflow-y: auto; white-space: pre-wrap; word-break: break-all; color: #475569;"><%=Trim(Replace(arrData(11,i) & "", "자동감지/", ""))%></div>
                                                <% If arrData(10,i) <> "" Then %>
                                                    <div style="margin-top:8px; border-top:1px dashed #cbd5e1; padding-top:6px; color:#64748b; font-size:10px;">
                                                        <strong>대상 키:</strong> <%=arrData(10,i)%>
                                                    </div>
                                                <% End If %>
                                            </div>
                                        </div>
                                    </td>
                                </tr>
                                <%          
                                            vNum = vNum - 1
                                        Next 
                                %>
                                <% End If %>
                            </tbody>
                        </table>
                        
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
