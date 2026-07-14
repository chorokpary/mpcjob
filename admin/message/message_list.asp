<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : message_list.asp / 쪽지 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strTabFlag

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 20	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspNote_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@RecvSeq").Value 	= Session("ASeq")
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), Seq(1), RecrSeq(2), SendName(3), IsAnswer(4), RegDate(5), Memo(6)
			kFields 	= Array("TotalCnt","Seq","RecrSeq","SendName","IsAnswer","RegDate","Memo")
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
	<title>쪽지 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 쪽지가 없습니다.");
			}else{
				if(confirm("선택한 쪽지를 삭제하시겠습니까?")){
					$("#Flag").val("DEL_RECV");
					$("#frmList").attr({action:"message_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function goForm(seq){
			goViewCommon(seq,"message_view.asp");
		}
		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
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
					<h2>쪽지함</h2>
					<div class="searchArea">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Seq" name="Seq">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<select id="sKey" name="sKey" title="검색항목" class="w80">
								<option value="">검색선택</option>
								<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>이름</option>
								<option value="CONTENT"<%=Fn_SetDefault(sKey,"CONTENT"," selected","")%>>내용</option>
							</select>
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>" class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
						</form>
					</div>
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#" id="btnChkAll">전체선택</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnDel">선택삭제</a></span>
						</div>
					</div>
					
					<div class="tblArea">
						<form name="frmList" id="frmList">
						<input type="hidden" id="Flag" name="Flag">
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
									<th>발송자</th>
									<th>내용</th>
									<th>작성일자</th>
									<th>확인</th>
									<th>관리</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="7">등록된 쪽지가 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), Seq(1), RecrSeq(2), SendName(3), IsAnswer(4), RegDate(5), Memo(6)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td><%=arrData(3,i)%></td>
									<td class="alignL"><a href="javascript:goForm('<%=arrData(1,i)%>');"><%=FN_StringCut(HTMLEncode(arrData(6,i)), 1, 30,"...")%></a><%=Fn_SetDefault(arrData(4,i),"True"," <img src=""/admin/images/ico_reply.gif"" align=""absmiddle"">","")%></td>
									<td><%=FN_SetDateTimeFormat(arrData(5,i),"YYYY.MM.DD")%></td>
									<td>
										<% If arrData(2,i) <> "0" Then %>
										<span class="btns btn_b"><a href="/recruit/recruit_view.asp?Seq=<%=arrData(2,i)%>" target="_blank">공고</a></span>
										<% End If %>
									</td>
									<td>
										<span class="btns btn_in"><a href="javascript:goForm('<%=arrData(1,i)%>');">답변</a>
										<% If arrData(2,i) <> "0" Then %>
										</span><span class="btns btn_in"><a href="/admin/recruit/job_skin_mod.asp?Flag=MOD&Seq=<%=arrData(2,i)%>">공고수정</a></span>
										<% Else %>
										</span>
										<% End If %>
									</td>
								</tr>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
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