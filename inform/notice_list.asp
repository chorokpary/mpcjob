<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : notice_list.asp / 공지사항 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
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

		listPage 	= 5
		listRecord 	= 10	'목록갯수
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
			.Parameters("@Flag").Value 		= "FRT"
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
	<title>HC알리미 - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function goView(seq){
			goViewCommon(seq, "notice_view.asp");
		}
	</script>
</head>
<body>
	<div id="wrap">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_notice.jpg" alt="공지사항" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="board">
					<table>
						<caption>공지사항 리스트</caption>
						<colgroup>
							<col class="w64" />
							<col />
							<col class="w79" />
							<col class="w100" />
							<col class="w78" />
						</colgroup>
						<thead>
							<tr>
								<th>번호</th>
								<th>제목</th>
								<th>작성자</th>
								<th>등록일</th>
								<th>조회수</th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="5">등록된 글이 없습니다.</td>
							</tr>
							<% Else %>
							<% 		
									' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6)
									vNum = RecordCount - ((Page-1) * listRecord) 
									For i = 0 To arrDataNum 
							%>
							<tr>
								<td><%=vNum%></td>
								<td class="alignL"><a href="javascript:goView('<%=arrData(1,i)%>');"><%=FN_ClearTag(arrData(2,i))%></a></td>
								<td><%=arrData(4,i)%></td>
								<td><%=FN_SetDateTimeFormat(arrData(3,i),"YYYY-MM-DD")%></td>
								<td><%=arrData(5,i)%></td>
							</tr>
							<%			vNum = vNum - 1
									Next 
							%>
							<% End If %>
						</tbody>
					</table>
				</div>
				<div class="paging">
					<% Call SB_PagingWrite(listPage, RecordCount, PageCount) %>
				</div>

				<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				<input type="hidden" id="Seq" name="Seq">
				<div class="searchArea">
					<select  id="sKey" name="sKey" title="검색항목" class="w80">
						<option value="ALL"<%=Fn_SetDefault(sKey,"ALL"," selected","")%>>전체</option>
						<option value="TITLE"<%=Fn_SetDefault(sKey,"TITLE"," selected","")%>>제목</option>
						<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>작성자</option>
					</select>
					<input type="text" id="sWord" name="sWord" title="검색어" value="<%=sWord%>" class="w190" />
					<input type="image" alt="검색" src="/images/common/btn_search.jpg">
				</div>
				</form>

			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>