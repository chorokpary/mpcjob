<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
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

		listPage 	= 5
		listRecord 	= 15	'목록갯수
		Page 		= FN_Req("Page","1")

		sCate 	= FN_Req("sCate","채용관련")
		
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
			.Parameters("@BbsKey").Value 	= "faq"
			.Parameters("@sCate").Value 	= sCate
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6), Contents(7)
			kFields 	= Array("TotalCnt","BbsSeq","Subject","RegDate","RegName","Hit","Category","Contents")
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
	<title>HC알리미 FAQ - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function goTab(cate){
			$("#sCate").val(cate);
			goPage(1);
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
						<h2><img src="/images/pages/tit_faq.jpg" alt="FAQ" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="tabArea">
					<span class="tab<%=Fn_SetDefault(sCate,"채용관련"," on","")%>"><a href="javascript:goTab('채용관련');">채용관련문의</a>
					</span><span class="tab<%=Fn_SetDefault(sCate,"이력서관련"," on","")%>"><a href="javascript:goTab('이력서관련');">이력서관련문의</a></span>
				</div>
				<div class="faq">
					<% If Not kData Then %>
					<dl>
						<dt>등록된 글이 없습니다.</dt>
					</dl>
					<tr>
						<td colspan="5">등록된 글이 없습니다.</td>
					</tr>
					<% Else %>
					<% 		
							' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6), Contents(7)
							vNum = RecordCount - ((Page-1) * listRecord) 
							For i = 0 To arrDataNum 
					%>
					<dl>
						<dt><%=FN_ClearTag(arrData(2,i))%></dt>
						<dd><div><%=FN_NLToBr(arrData(7,i))%></div></dd>
					</dl>
					<%			vNum = vNum - 1
							Next 
					%>
					<% End If %>
				</div>
				<div class="paging">
					<% Call SB_PagingWrite(listPage, RecordCount, PageCount) %>
				</div>
				
				<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="Seq" name="Seq">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				<input type="hidden" id="sCate" name="sCate" value="<%=sCate%>">
				</form>

			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>