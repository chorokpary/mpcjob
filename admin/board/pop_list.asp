<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pop_list.asp / 팝업관리 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-22 / 강미현 / 4M / 최초작성
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
			.CommandText ="uspConfPopup_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), Seq(1), Title(2), IsView(3), BDate(4), EDate(5)
			kFields 	= Array("TotalCnt","Seq","Title","IsView","BDate","EDate")
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
	<title>팝업관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 팝업이 없습니다.");
			}else{
				if(confirm("선택한 팝업을 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"pop_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function goForm(seq){
			window.location.href="pop_form.asp?Page=<%=Page%>&Seq="+seq;
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
					<h2>팝업관리</h2>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">팝업등록</a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<select id="sKey" name="sKey" title="검색항목">
								<option value="TITLE"<%=Fn_SetDefault(sKey,"TITLE"," selected","")%>>제목</option>
								<option value="CONTENT"<%=Fn_SetDefault(sKey,"CONTENT"," selected","")%>>내용</option>
							</select>
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>" class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
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
								<col class="w65" />
								<col />
								<col class="w160" />
							</colgroup>
							<thead>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th>공개</th>
									<th>제목</th>
									<th>기간</th>
									
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="5">등록된 팝업이 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), Seq(1), Title(2), IsView(3), BDate(4), EDate(5)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td><%=Fn_SetDefault(arrData(3,i),"True","공개","비공개")%></td>
									<td class="alignL"><a href="javascript:goForm('<%=arrData(1,i)%>');"><%=FN_ClearTag(arrData(2,i))%></a></td>
									<td><%=arrData(4,i)%> ~ <%=arrData(5,i)%></td>
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