<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : apply.asp / 입사지원 내역
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, Seq
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim StrUsrName
	
	Call SB_LoginChkPage()
	Call Main()
	
	Sub Main()

		listPage 	= 5
		listRecord 	= 10	'목록갯수
		Page 		= FN_Req("Page","1")


		Seq 	= Session("USeq")
		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		StrUsrName = getName(Seq)
		
	End Sub

	Function getName(argSeq)

		Dim Rs, Rtn
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "SIMP"
			.Parameters("@Seq").Value 	= Seq
			
			Set Rs = .Execute
		End With
			
		If Not Rs.Eof Then
			Rtn = Rs("UsrName")
		Else
			Rtn = ""
		End If
		
		Rs.Close
		Set Rs = Nothing
		
		getName = Rtn

	End Function
		
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@Seq").Value 		= Seq
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), AppSeq(1), RecrSeq(2), RegDate(3), Status(4), PrjName(5), Title(6), AppBDate(7), AppEDate(8), IsView(9)(Recruit), CntQues(10)
			kFields 	= Array("TotalCnt","AppSeq","RecrSeq","RegDate","Status","PrjName","Title","AppBDate","AppEDate","IsView", "CntQues")
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
	<title>MY PAGE - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function doDel(seq){
			$("#frmPage").find("#Seq").val(seq);
			$("#frmPage").attr({action:"apply_delete.asp", method:"POST"}).submit();
		}
		
		function applyCancel(seq){
			openLayer("apply_confirm.asp","Seq=" + seq, "w668");
		}
		
		function applyQuestion(seq){
			openLayer("/mypage/apply_question_ajax.asp","RecrSeq="+seq, "w668");
			//alert("사전질문 수정 (" +  seq + ")");
		}
	</script>
</head>
<body>
	<div id="wrap" class="applyMng">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_applyMng.jpg" alt="지원관리" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="section">
					<table class="cell_alignC">
						<caption>이력서 개인정보</caption>
						<colgroup>
							<col class="w185" />
							<col />
							<col class="w85" />
							<col class="w85" />
							<col class="w165" />
						</colgroup>
						<thead>
							<tr>
								<th><label for="">채용팀</label></th>
								<th><label for="">채용제목</label></th>
								<th><label for="">상태</label></th>
								<th><label for="">마감일</label></th>
								<th><label for="">입사지원</label></th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="5">지원 내역이 없습니다.</td>
							</tr>
							<% Else %>
							<% 		
									' TotalCnt(0), AppSeq(1), RecrSeq(2), RegDate(3), Status(4), PrjName(5), Title(6), AppBDate(7), AppEDate(8), IsView(9), CntQues(10)
									vNum = RecordCount - ((Page-1) * listRecord) 
									For i = 0 To arrDataNum 
							%>
							<tr>
								<td><%=arrData(5,i)%></td>
								<td><%=arrData(6,i)%></td>
								<td>
									<% If FN_SetDateTimeFormat(arrData(8,i),"YYYY-MM-DD") >= Date() And arrData(9,i) Then %>
									<img src="/images/pages/ico_accepting.jpg" alt="접수중" />
									<% Else %>
									<img src="/images/pages/ico_termination.jpg" alt="접수종료" />
									<% End If %>
								</td>
								<td><%=FN_SetDateTimeFormat(arrData(8,i),"YYYY/MM/DD")%></td>
								<td class="alignL" style="padding-left:5px;">
									<% If FN_SetDateTimeFormat(arrData(8,i),"YYYY-MM-DD") >= Date() And arrData(9,i) Then %>
										<span class="btns_s"><a href="javascript:applyCancel(<%=arrData(1,i)%>);" class="small">지원취소</a></span>
										<% If arrData(10,i) > 0 Then %>
										<span class="btns_s"><a href="javascript:applyQuestion(<%=arrData(2,i)%>)" class="small">사전질문수정</a></span>
										<% End If %>
									<% End If %>
								</td>
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
				</form>

			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>