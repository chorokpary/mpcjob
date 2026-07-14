<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : notice_list.asp / 공지사항 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-29 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 15	'목록갯수
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
			.CommandText ="uspBbs_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@BbsKey").Value 	= "notice"
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6)
			kFields 	= Array("TotalCnt","BbsSeq","Subject","RegDate","RegName","Hit","Category")
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
	<title>공지사항 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 글이 없습니다.");
			}else{
				if(confirm("선택한 글을 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"notice_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function goForm(seq){
			window.location.href="notice_form.asp?Page=<%=Page%>&Seq="+seq;
		}
		
		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
		    $("#btnReg").click(function(){ goForm(0);});
		})
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
					<h2>공지사항</h2>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">글쓰기</a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<select id="sKey" name="sKey" title="검색항목">
								<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>작성자</option>
								<option value="TITLE"<%=Fn_SetDefault(sKey,"TITLE"," selected","")%>>제목</option>
								<!--option value="CONTENT"<%=Fn_SetDefault(sKey,"CONTENT"," selected","")%>>내용</option//-->
							</select>
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>"  class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
						</form>
						</div>
					</div>
					
					<div class="tblArea">
						<form name="frmList" id="frmList">
						<input type="hidden" id="Flag" name="Flag">
						<table>
							<colgroup>
								<col class="w40" />
								<col class="w40" />
								<col />
								<col class="w90" />
								<col class="w85" />
								<col class="w50" />
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
									<th>선택</th>
									<th>번호</th>
									<th>제목</th>
									<th>등록일</th>
									<th>작성자</th>
									<th>조회수</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="6">등록된 글이 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td class="alignL"><a href="javascript:goForm('<%=arrData(1,i)%>');"><%=FN_ClearTag(arrData(2,i))%></a></td>
									<td><%=FN_SetDateTimeFormat(arrData(3,i),"YYYY-MM-DD")%></td>
									<td><%=arrData(4,i)%></td>
									<td><%=arrData(5,i)%></td>
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