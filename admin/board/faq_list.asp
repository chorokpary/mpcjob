<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : faq_list.asp / FAQ 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-02 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, sCate
	Dim arrData, arrDataNum, kData, kFields, i, vNum

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 15	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		sCate 	= FN_Req("sCate","")
		
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
			.Parameters("@BbsKey").Value 	= "faq"
			.Parameters("@sCate").Value 	= sCate
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
	<title>FAQ - MPC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 글이 없습니다.");
			}else{
				if(confirm("선택한 글을 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"faq_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function openForm(seq){
			openWin("faq_form.asp?Seq="+seq,"popFaq",0,0,930,650,0)
		}
		
		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
		    $("#btnReg").click(function(){ openForm(0);});
		    
		    $("#sCate").change(function(){	$("#frmPage").submit();	});
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
					<h2>FAQ</h2>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">글쓰기</a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_BBS", "Category_faq", "sCate", "전체분류", "", sCate, "등록된 카테고리가 없습니다.", " ", False)%>
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
								<col class="w110" />
								<col />
							</colgroup>
							<thead>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th>분류</th>
									<th>질문</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="4">등록된 글이 없습니다.</td>
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
									<td><%=arrData(6,i)%></td>
									<td class="alignL"><a href="javascript:openForm('<%=arrData(1,i)%>');"><%=FN_ClearTag(arrData(2,i))%></a></td>
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