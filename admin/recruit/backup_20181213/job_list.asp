<%
'____________________________________________________________________________________
'
' * Discription : job_list.asp / 채용공고 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim sViewFlag, sBDate, sEDate
	Dim StrViewFlag : StrViewFlag = ""

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= FN_Req("ListSize","20")	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))

		sViewFlag = FN_Req("sViewFlag","")
		sBDate 	= FN_Req("sBDate","")
		sEDate 	= FN_Req("sEDate","")

		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 20
		End If
		

		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord

			.Parameters("@sIsView").Value 	= CONF_ISVIEW
			.Parameters("@sViewFlag").Value = sViewFlag
			.Parameters("@sBDate").Value 	= sBDate
			.Parameters("@sEDate").Value 	= sEDate
			
			.Parameters("@AdmLevel").Value 	= Session("ALevel")
			.Parameters("@AdmSeq").Value 	= Session("ASeq")

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), RecrSeq(1), Title(2), IsMain(3), IsTop(4), AppBDate(5), AppEDate(6)
			' PrjName(7), AdmName(8), AppCnt(9), IsView(10)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","IsMain","IsTop","AppBDate","AppEDate","PrjName","AdmName","AppCnt","IsVIew")
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
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript">
		function chkApp(){
			if (cntChecked("#frmList")==0){
				alert("선택된 채용공고가 없습니다.");
			}else{
				openWin("job_apply_stats.asp?Seq="+ getCheckedVal("Seq"),"popApplyStats",0,0,930,770,0);
			}
		}

		function chkStat(){
			if (cntChecked("#frmList")==0){
				alert("선택된 채용공고가 없습니다.");
			}else{
				openWin("job_partner_stats.asp?Seq="+ getCheckedVal("Seq"),"popPartnerStats",0,0,930,770,0);
			}
		}
		
		function doStatus(seq){
			$("#frmList").find("#Flag").val("FIN");
			$("#frmList").find("#Seq_" + seq).attr("checked",true);
			$("#frmList").attr({action:"job_proc.asp?Seq=" + seq , method:"post"}).submit();
		}
		
		function doStatusFin(){
			alert("이미 마감된 공고입니다.");
		}
		
		function goFront(seq){
			var win = window.open("about:blank");
			win.location.href="/recruit/recruit_view.asp?Seq="+seq;
		}
	
		function goForm(seq, flag){
			$("#frmPage").find("#Seq").val(seq);
			$("#frmPage").find("#Flag").val(flag);
			$("#frmPage").attr({action:"job_skin_mod.asp", method:"post"}).submit();
			//window.location.href="job_skin_mod.asp?Page=<%=Page%>&Seq="+seq + "&Flag=" + flag;
		}

		function goApply(seq){
			// 프로젝트 단위 새창 오픈
			//var win = 	openWin("job_apply_user.asp?RecrSeq="+ seq,"popApplyUser_"+seq,0,0,1020,770,0);
			// 클릭수 단위 새창 오픈
			var win = 	openWin("job_apply_user.asp?RecrSeq="+ seq,"",0,0,1020,770,0);
			win.focus();
		}
		
	
		
		$(function() {
		    $("#btnChkAll").click(function(){ chekAllBtn("#frmList","#btnChkAll"); });
		    $("#btnStat").click(chkStat);
		    $("#btnApp").click(chkApp);
		    $("#btnRecom").click(function(){ openWin("job_recom_list.asp","popRecom",0,0,930,630,0); });

		    $("#ListSize").change(function(){ goPage(1); });
		    $("#sViewFlag").change(initPage);
		    $("#sBDate").change(initPage);
		    $("#sEDate").change(initPage);
		    

			$("#sBDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#sEDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});

		});
	</script>
					<form name="frmPage" id="frmPage" method="post">
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="Seq" name="Seq">
					<input type="hidden" id="Flag" name="Flag">
					<div class="searchArea">
						<div class="search_date">
							<select id="sViewFlag" name="sViewFlag" title="노출 필터링" class="w110">
								<option value="">전체</option>
								<option value="TOP"<%=Fn_SetDefault(sViewFlag,"TOP"," selected","")%>>스페셜 HOT</option>
								<option value="MAIN"<%=Fn_SetDefault(sViewFlag,"MAIN"," selected","")%>>이달의 추천공고</option>
							</select>
							<input type="text" id="sBDate" name="sBDate" value="<%=sBDate%>" class="w80" style="margin-right:5px;" maxlength="10" /> 부터 ~ 
							<input type="text" id="sEDate" name="sEDate" value="<%=sEDate%>" class="w80" style="margin-right:5px;" maxlength="10" /> 까지
							<!--select name="sKey" id="sKey" title="검색선택" class="w80">
								<option value="">검색선택</option>
								<option value="CHARGE"<%=Fn_SetDefault(sKey,"CHARGE"," selected","")%>>담당자</option>
								<option value="PROJECT"<%=Fn_SetDefault(sKey,"PROJECT"," selected","")%>>프로젝트명</option>
								<option value="TITLE"<%=Fn_SetDefault(sKey,"TITLE"," selected","")%>>공고제목</option>
							</select//-->
							/ 통합검색 : <input type="text" id="sWord" name="sWord" value="<%=sWord%>" title="검색어" class="w190" />
							<span class="btns btn_b"><button type="submit">검색</button></span>
							<p><span class="ess">*</span> 검색일자의 기준은 프로젝트 마감일을 기준으로 합니다.</p>
						</div>
					</div>
					<div class="btnArea">
						<div class="fl_l">
							<select name="ListSize" id="ListSize" title="출력개시">
								<option value="20">출력개수</option>
								<option value="5"<%=Fn_SetDefault(listRecord,"5"," selected","")%>>5개</option>
								<option value="10"<%=Fn_SetDefault(listRecord,"10"," selected","")%>>10개</option>
								<option value="20"<%=Fn_SetDefault(listRecord,"20"," selected","")%>>20개</option>
								<option value="30"<%=Fn_SetDefault(listRecord,"30"," selected","")%>>30개</option>
								<option value="50"<%=Fn_SetDefault(listRecord,"50"," selected","")%>>50개</option>
								<option value="100"<%=Fn_SetDefault(listRecord,"100"," selected","")%>>100개</option>
							</select>
							<span class="btns btn_g"><a href="#" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_b"><a href="#" id="btnStat">통계보기</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnApp">지원자현황</a>
							</span><% If Session("ALevel") = "1" Then %><span class="btns btn_b"><a href="job_skin_reg.asp">채용공고등록</a>
							</span><span class="btns btn_b"><a href="#" id="btnRecom">추천공고 관리</a></span><% End If %>
						</div>
					</div>
					</form>

					<form name="frmList" id="frmList">
					<input type="hidden" id="Flag" name="Flag">
					<div class="tblArea">
						<table>
							<colgroup>
								<col class="w40" />
								<col class="w40" />
								<col class="w80" />
								<col />
								<col class="w85" />
								<col class="w60" />
								<col class="w120" />
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
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th>담당자</th>
									<th>프로젝트명</th>
									<th>노출</th>
									<th>총 지원자</th>
									<th>관리</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="7">등록된 채용공고가 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), RecrSeq(1), Title(2), IsMain(3), IsTop(4), AppBDate(5), AppEDate(6)
										' PrjName(7), AdmName(8), AppCnt(9)
										vNum = RecordCount - ((Page-1) * listRecord) 
										
										For i = 0 To arrDataNum 
											If arrData(3,i) Or arrData(4,i) Then
												If arrData(3,i) Then
													StrViewFlag = "이달의 추천공고"
												End If
												StrViewFlag = Fn_SetDefault(arrData(3,i),True,"이달의 <br>추천공고","")
												
												If arrData(4,i) Then
													If StrViewFlag <> "" Then 		StrViewFlag = StrViewFlag & ",<br>"
													StrViewFlag = StrViewFlag & "스페셜 HOT"
												End If
											Else
												StrViewFlag = "일반채용공고"
											End If
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td><%=arrData(8,i)%></td>
									<td class="alignL">
										<div class="job_info">
											<strong>
												<% If CONF_ISVIEW = "-1" Then %><%=Fn_SetDefault(arrData(10,i),"False","<span class='red'>[마감] </span>","<span class='blue'>[진행중]</span> ") %><% End If %>
												<%=arrData(7,i)%>
											</strong>
											<a href="javascript:goForm(<%=arrData(1,i)%>,'MOD');"><%=arrData(2,i)%></a>
										</div>
										<div class="job_detail">
											<p>&middot; 전형시작일 : <%=FN_SetDateTimeFormat(arrData(5,i),"YYYY-MM-DD")%></p>
											<p>&middot; 전형마감일 : <%=FN_SetDateTimeFormat(arrData(6,i),"YYYY-MM-DD")%></p>
										</div>
									</td>
									<td><%=StrViewFlag %></td>
									<td><%=arrData(9,i)%><br /><span class="btns btn_in"><a href="javascript:goApply(<%=arrData(1,i)%>);">보기</a></span>
									</td>
									<td>
										<span class="btns btn_in"><a href="javascript:goForm(<%=arrData(1,i)%>,'MOD');">수정</a>
										</span><span class="btns btn_in"><a href="javascript:goFront(<%=arrData(1,i)%>);">공고보기</a></span>
										<span class="btns btn_in"><a href="javascript:goForm(<%=arrData(1,i)%>,'REF');" style="width:48px;">재등록</a>
										</span><span class="btns btn_in"><a href="javascript:<%=Fn_SetDefault(arrData(10,i),"True","doStatus("&arrData(1,i)&")","doStatusFin()") %>;" >마감</a></span>
									</td>
								</tr>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
					</div>
					</form>

					<div class="paging">
						<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
					</div>
