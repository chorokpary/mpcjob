<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_apply_user.asp / 지원자 관리 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.Buffer = True '버퍼 사용

	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim sStatus, sBDate, sEDate
	Dim SortKey, SortDIr
	Dim RecrSeq
	Dim StrTitle : StrTitle = ""
	Dim CntQues : CntQues = 0
	Dim strAppCStyle, strAppCMsg

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= FN_Req("ListSize","100")	'목록갯수
		Page 		= FN_Req("Page","1")

		RecrSeq 	= FN_Req("RecrSeq","0")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))

		sStatus = FN_Req("sStatus","")
		sBDate 	= FN_Req("sBDate","")
		sEDate 	= FN_Req("sEDate","")

		SortKey = FN_Req("SortKey","")
		SortDIr = FN_Req("SortDIr","DESC")

		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		
		If sStatus  = "All" Then 	sStatus = ""
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 100
		End If
		
		If Session("ALevel") <> "1" Then	' 최종관리자가 아닐경우 접근 권한 체크
			Call chkAuth()
		End If
		
		Call getRecruitInfo()
		Call getList()
		
	End Sub
	
	Sub chkAuth()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruitCharge_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "AUTH"
			.Parameters("@RecrSeq").Value 	= RecrSeq
			.Parameters("@AdmSeq").Value 	= Session("ASeq")

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			'권한 없음
			Call SB_ReturnErr("열람 권한이 없는 공고입니다.","CLOSE")
		End If
	
		Rs.Close
		Set Rs = Nothing

	End Sub

	Sub getRecruitInfo()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "INFO"
			.Parameters("@Seq").Value 		= RecrSeq

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			'권한 없음
			Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","CLOSE")
		Else
			StrTitle = "[" & Rs("PrjName") & "] - " & Rs("Title")
			CntQues = Rs("CntQues")
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub	
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM_R"
			.Parameters("@Seq").Value 		= RecrSeq
			
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= Fn_SetDefault(sKey,"MOBILE",Replace(sWord,"-",""),sWord) 

			.Parameters("@sStatus").Value 	= sStatus
			.Parameters("@sBDate").Value 	= sBDate
			.Parameters("@sEDate").Value 	= sEDate
			If SortKey <> "" And SortDir = "DESC" Then
				.Parameters("@SortDesc").Value 	= SortKey
			ElseIf SortKey <> "" And SortDir = "ASC" Then
				.Parameters("@SortAsc").Value 	= SortKey
			End If

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), AppSeq(1), UsrSeq(2), PartnerDepth1(3), PartnerDepth2(4), RegDate(5)
			' UsrName(6), Gender(7), Birthday(8), AddrSido(9), AddrGugun(10)
			' Attach1_Path(11), Attach1_FName(12), Attach2_Path(13), Attach2_FName(14), Attach3_Path(15), Attach3_FName(16)
			' CntMemo(17), Status(18), CntAll(19), CntIng(20), CntFin(21), UsrType(22), Mobile(23), IsBlack(24), CancelDate(25), IsOut(26)
			kFields 	= Array("TotalCnt","AppSeq","UsrSeq","PartnerDepth1","PartnerDepth2","RegDate","UsrName","Gender","Birthday","AddrSido","AddrGugun" _
							,"Attach1_Path","Attach1_FName","Attach2_Path","Attach2_FName","Attach3_Path","Attach3_FName","CntMemo","Status","CntAll","CntIng","CntFin","UsrType","Mobile","IsBlack","CancelDate","IsOut")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True

			RecordCount = arrData(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>지원자 관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원을 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"job_apply_proc.asp", method:"post", target:"_self"}).submit();
				}
			}
		}

		function doChangeStatus(status){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원의 상태를 변경하시겠습니까?")){
					$("#Flag").val("STA");
					$("#Status").val(status);
					$("#frmList").attr({action:"job_apply_proc.asp", method:"post", target:"_self"}).submit();
				}
			}
		}

		function doSendSMS(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				openWin("/admin/popup/sms.asp?Flag=RECR&Seq="+ getCheckedVal("Seq") ,"popSMS",0,0,948,550,1);
			}
		}

		function doDownXls(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원을 엑셀로 다운 받으시겠습니까?")){
					$("#frmList").attr({action:"job_apply_user_xls.asp", method:"post", target:"_self"}).submit();
				}
			}
		}

		function goSort(key,dir){
			$("#SortKey").val(key);
			$("#SortDir").val(dir);
			goPage(<%=Page%>);
		}
		
		function goUserInfo(seq){ 
			openWin("../member/mem_print.asp?RecrSeq=<%=RecrSeq%>&Seq="+ seq,"popPrint",0,0,930,650,1);
		}

		function goUserMemo(seq){ 
			openWin("../member/mem_memo.asp?Seq="+ seq,"popMemo",0,0,930,620,0);
		}

		function goUserApply(seq){ 
			openWin("../member/mem_apply.asp?Seq="+ seq,"popApply",0,0,930,550,0);
		}
		
		function goUserQuestion(seq){
			openWin("../member/mem_question.asp?RecrSeq=<%=RecrSeq%>&Seq="+ seq,"popQues",0,0,930,650,1);
		}
		
		
		$(function() {
			$("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
			$("#btnDel").click(chkDel);
			$("#btnSMS").click(doSendSMS);
			$("#btnXls").click(doDownXls);

			$("#ListSize").change(function(){ goPage(1); });
			$("#sStatus").change(function(){ goPage(1); });
		});
	</script>
</head>

<body class="pop w1020">
	<div id="wrap" class="applicant recruitManage">
		<div id="contArea">
			<h2>지원자 관리 (<%=StrTitle%>)</h2>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('00');">접수완료</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('01');">서류전형합격</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('02');">서류전형불합격</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('03');">면접불참석</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('10');">결과대기</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('04');">최종합격</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('05');">최종불합격</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('06');">교육참석</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('07');">교육미참석</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('09');">교육이탈</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('08');">입사</a></span>
					<span class="btns btn_b2"><a href="javascript:doChangeStatus('11');">채용취소</a></span>
				</div>
			</div>
			<form name="frmPage" id="frmPage" method="post" action="job_apply_user.asp">
			<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
			<input type="hidden" id="Page" name="Page" value="<%=Page%>">
			<input type="hidden" id="SortKey" name="SortKey" value="<%=SortKey%>">
			<input type="hidden" id="SortDir" name="SortDir" value="<%=SortDir%>">
			<div class="btnArea">
				<div class="fl_l">
					<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_APPLY", "Status", "sStatus", "진행단계", "", sStatus, "", "", False)%>
					<select name="sKey" id="sKey" title="검색설정" class="w80">
						<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>이름</option>
						<option value="MOBILE"<%=Fn_SetDefault(sKey,"MOBILE"," selected","")%>>핸드폰번호</option>
					</select>
					<input type="text" name="sWord" id="sWord" value="<%=sWord%>" title="검색어" class="w200" />

					<span class="btns btn_b"><button type="submit">검색</button>
					</span><span class="btns btn_b"><a href="#btnXls" id="btnXls">엑셀저장</a></span>
				</div>
				<div class="fl_r">
					<select name="ListSize" id="ListSize" title="출력개시">
						<option value="5"<%=Fn_SetDefault(listRecord,"5"," selected","")%>>5개</option>
						<option value="10"<%=Fn_SetDefault(listRecord,"10"," selected","")%>>10개</option>
						<option value="20"<%=Fn_SetDefault(listRecord,"20"," selected","")%>>20개</option>
						<option value="30"<%=Fn_SetDefault(listRecord,"30"," selected","")%>>30개</option>
						<option value="50"<%=Fn_SetDefault(listRecord,"50"," selected","")%>>50개</option>
						<option value="100"<%=Fn_SetDefault(listRecord,"100"," selected","")%>>100개</option>
						<option value="-1"<%=Fn_SetDefault(listRecord,"-1"," selected","")%>>전체보기</option>
					</select>
					<span class="btns btn_b"><a href="#btnSMS" id="btnSMS">결과통보</a></span>
				</div>
			</div>
			</form>
			<% Response.Flush %>
			<div class="inner">
				<form name="frmList" id="frmList">
				<input type="hidden" id="Flag" name="Flag">
				<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
				<input type="hidden" id="Status" name="Status">
				<input type="hidden" id="xStatus" name="xStatus" value="<%=sStatus%>">
				<input type="hidden" id="xKey" name="xKey" value="<%=sKey%>">
				<input type="hidden" id="xWord" name="xWord" value="<%=sWord%>">
				<input type="hidden" id="xListSize" name="xListSize" value="<%=listRecord%>">
				<input type="hidden" id="SortXlsKey" name="SortXlsKey" value="<%=SortKey%>">
				<input type="hidden" id="SortXlsDir" name="SortXlsDir" value="<%=SortDir%>">
				<div class="tblArea">
					<table>
						<colgroup>
							<col class="w30" />
							<col class="w40" />
							<col class="w50" />
							<col class="w60" />
							<col class="w35" />
							<col class="w35" />
							<col />
							<col class="w90" />
							<col class="w95" />
							<col class="w150" />
							<col class="w30" />
							<col class="w135" />
							<col class="w70" />
							<col class="w40" />
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
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
							</tr>
							<tr>
								<th>선택</th>
								<th>번호</th>
								<th>이름</th>
								<th>첨부</th>
								<th>
									<% If SortKey = "GENDER" And SortDir = "ASC" Then %>
									성별<a href="javascript:goSort('GENDER','DESC')">▲</a>
									<% Else %>
									성별<a href="javascript:goSort('GENDER','ASC')">▼</a>
									<% End If %>
								</th>
								<th>
									<% If SortKey = "AGE" And SortDir = "ASC" Then %>
									나이<a href="javascript:goSort('AGE','DESC')">▲</a>
									<% Else %>
									나이<a href="javascript:goSort('AGE','ASC')">▼</a>
									<% End If %>
								</th>
								<th>거주지</th>
								<th>전화번호</th>
								<th>상태</th>
								<th>지난지원현황</th>
								<th>메모</th>
								<th>지원경로</th>
								<th>지원일</th>
								<th>사전<br>질문</th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="14">검색된 지원자가 없습니다.</td>
							</tr>
							<% Else %>
							<% 		
									' TotalCnt(0), AppSeq(1), UsrSeq(2), PartnerDepth1(3), PartnerDepth2(4), RegDate(5)
									' UsrName(6), Gender(7), Birthday(8), AddrSido(9), AddrGugun(10)
									' Attach1_Path(11), Attach1_FName(12), Attach2_Path(13), Attach2_FName(14), Attach3_Path(15), Attach3_FName(16)
									' CntMemo(17), Status(18), CntAll(19), CntIng(20), CntFin(21), UsrType(22), Mobile(23), IsBlack(24), CancelDate(25), IsOut(26)
									
									vNum = RecordCount - ((Page-1) * listRecord) 

									For i = 0 To arrDataNum 
									
										strAppCStyle = "" : strAppCMsg = ""

										If Not IsNull(arrData(25,i)) Then
											strAppCStyle = "cancel"
											strAppCMsg = "<br>/ 지원취소"
										End If

										If  arrData(26,i) Then
											strAppCStyle = "cancel"
											strAppCMsg = "<br>/ 탈퇴"
										End If
										
										If arrData(24,i) Then
											strAppCStyle = strAppCStyle & " blackList"
										End If
										
										If strAppCStyle <> "" Then
											strAppCStyle = " class=""" & strAppCStyle & """"
										End If
							%>
							<tr <%=strAppCStyle%>>
								<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
								<td><%=vNum%></td>
								<td><a href="javascript:goUserInfo('<%=arrData(2,i)%>');"><%=arrData(6,i)%></a>
									<% If Trim(arrData(22,i)) = "0" Then %><img src="/admin/images/ico_easy.jpg" alt="간편가입회원" /><% End If %></td>
								<td>
									<% If Trim(arrData(11,i)) <> "" Then %><a href="javascript:goDown('<%=arrData(2,i)%>','USERS','Attach1_Path','<%=arrData(12,i)%>');" title="<%=arrData(12,i)%>"><img src="/admin/images/ico_file.png" alt="" /></a><% End If %>
									<% If Trim(arrData(13,i)) <> "" Then %><a href="javascript:goDown('<%=arrData(2,i)%>','USERS','Attach2_Path','<%=arrData(14,i)%>');" title="<%=arrData(14,i)%>"><img src="/admin/images/ico_file.png" alt="" /></a><% End If %>
									<% If Trim(arrData(15,i)) <> "" Then %><a href="javascript:goDown('<%=arrData(2,i)%>','USERS','Attach3_Path','<%=arrData(16,i)%>');" title="<%=arrData(16,i)%>"><img src="/admin/images/ico_file.png" alt="" /></a><% End If %>
								</td>
								<td>
									<%
										Select Case arrData(7,i) 
											Case "1" : Response.Write "남" 
											Case "2" : Response.Write "여" 
											Case Else : Response.Write "-" 
										End Select 
									%>
								</td>
								<td><% If IsNumeric(Left(arrData(8,i),4)) Then %><%=Year(Date) - Left(arrData(8,i),4) + 1%><% Else %> - <% End If %></td>
								<td><%=arrData(9,i)%>&nbsp;<%=arrData(10,i)%></td>
								<td><%=arrData(23,i)%></td>
								<td><%=arrData(18,i)%><%=strAppCMsg%></td>
								<td><a href="javascript:goUserApply('<%=arrData(2,i)%>');">총<%=arrData(19,i)%>건</a> (마감:<%=arrData(21,i)%>, 진행:<%=arrData(20,i)%>)</td>
								<td><a href="javascript:goUserMemo('<%=arrData(2,i)%>');"><%=arrData(17,i)%></a></td>
								<td><%=arrData(3,i)%> 
									<% If Not FN_IsBlank(arrData(4,i)) Then %>(<%=arrData(4,i)%>)<% End If %></td>
								<td><%=FN_SetDateTimeFormat(arrData(5,i),"YYYY-MM-DD")%></td>
								<td>
									<% If CntQues > 0 And IsNull(arrData(25,i)) And Not arrData(26,i) Then %>
									<span class="btns btn_b" style="margin:auto;"><a href="javascript:goUserQuestion('<%=arrData(2,i)%>');"  style="padding:0 2px 0 1px; color:#ffffff;">수정</a></span>
									<% End If %>
								</td>
							</tr>
							<%			vNum = vNum - 1
									Next 
							%>
							<% End If %>
						</tbody>
					</table>
					<div class="paging">
						<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
					</div>
				</div>
				</form>
			</div>
			<div class="btnArea bt1 pt10">
				<div class="fl_l">
					<span class="btns btn_g"><a href="javascript:;" id="btnChkAll">전체선택</a>
					</span><span class="btns btn_g"><a href="javascript:;" id="btnDel">선택삭제</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="javascript:;" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
	<form id="frmDown" name="frmDown">
		<input type="hidden" id="dnSeq" name="dnSeq">
		<input type="hidden" id="dnFld" name="dnFld">
		<input type="hidden" id="dnFName" name="dnFName">
		<input type="hidden" id="dnFlag" name="dnFlag">
	</form>
</body>
</html>