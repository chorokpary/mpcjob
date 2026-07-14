<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_flash.asp / 플래시 관리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-22 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-23 / 강미현 / 4M / 목록 타입 변경 (등록 Log)
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim sViewFlag
	Dim StrTitle
	Dim IsData
	
	Call Main()
	
	Sub Main()
	
		listPage = 10
		listRecord = 4	'목록갯수
		Page 		= FN_ReqF("Page","1")

		'sKey 	= FN_ReqF("sKey","")
		'sWord 	= FN_InputVal(FN_ReqF("sWord",""))

		sViewFlag = FN_ReqF("sViewFlag","ING")

		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		If sViewFlag = "ING" Then
			StrTitle = "사용(대기)중인 "
		Else
			StrTitle = "종료된 "
		End If

		Call getList()
	
	End Sub
	
	Sub getList()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfFlash_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sViewFlag").Value = sViewFlag
			
			' ## 회원정보
			Set Rs = .Execute

			If Rs.Eof or Rs.Bof Then
				arrDataNum 	= -1
				kData 		= False
			Else
				'TotalCnt(0), Seq(1), ImgPath(2), LinkUrl(3), Sort(4), BDate(5), EDate(6)
				kFields 	= Array("TotalCnt","Seq","ImgPath","LinkUrl","Sort","BDate","EDate")
				arrData 	= Rs.GetRows(,,kFields)
				arrDataNum 	= UBound(arrData, 2)
				kData 		= True

				RecordCount = arrData(0,0)
				PageCount = FN_PageCount(RecordCount,listRecord)

			End If
			
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>플래시 관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		/*
		function doSubmit(){
			var i, j;
			var chkSame = false;
			var v1, v2;
			
			for (i=1; i<=4;i++){
				for (j=1; j<=4;j++){
					// 값이 비어 있지 않은 정렬만 체크
					if ($("#Sort"+i).val() && $("#Sort"+j).val()){
						v1 = $("#Sort"+i).val();
						v2 = $("#Sort"+j).val();
						
						v1 = v1.substring(v1.length-1);
						v2 = v2.substring(v2.length-1);
						
						if (i != j && v1 == v2){
							chkSame = true;
						}
					}
				}
			}
			
			if (chkSame) {
				alert("동일한 순서를 입력할 수 없습니다.");
				return;
			}
			
			$("#frmList").attr({action:"conf_flash_proc.asp", method:"post"}).submit();
		}
		*/
		function goForm(seq, flag){
			$("#frmPage").find("#Seq").val(seq);
			$("#frmPage").find("#Flag").val(flag);
			$("#frmPage").attr({action:"conf_flash_form.asp", method:"post"}).submit();
		}

		function changeSort(sort,sortcur){
			$("#frmList").find("#Flag").val("SORT");
			$("#ChSort").val(sort);
			$("#ChSortCur").val(sortcur);
			$("#frmList").attr({action:"conf_flash_proc.asp", method:"post"}).submit();
		}


		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 배너가 없습니다.");
			}else{
				if(confirm("선택한 배너를 삭제하시겠습니까? 삭제한 내용은 복원 할 수 없습니다.")){
					$("#frmList").find("#Flag").val("DEL");
					$("#frmList").attr({action:"conf_flash_proc.asp", method:"post"}).submit();
				}
			}
		}

		
		$(function() {
			$("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
			$("#btnDel").click(chkDel);
			$("#btnPreview").click(function(){ openWin("conf_flash_preview.asp","popConfFlash_Preview",0,0,641,421,0); });
		 });
	</script>
</head>

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<div class="tabArea">
				<span class="tab<%=Fn_SetDefault(sViewFlag,"ING"," on","")%>"><a href="conf_flash.asp?sViewFlag=ING">사용(대기) 중</a>
				</span><span class="tab<%=Fn_SetDefault(sViewFlag,"FIN"," on","")%>"><a href="conf_flash.asp?sViewFlag=FIN">종료</a></span>
			</div>

			<h2><%=StrTitle%> 플래시 배너</h2>
			<% If sViewFlag = "ING" Then %>
			<p class="note_c">* 노출 대상이 4개 이상일 경우 사용자 페이지에서는 목록 하단부터 순서대로 4개만 노출되게 됩니다.</p>
			<% End If %>
			<div class="inner">
				<form name="frmPage" id="frmPage" method="post">
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="Seq" name="Seq">
					<input type="hidden" id="Flag" name="Flag">
					<input type="hidden" id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
				</form>
					
				<form id="frmList" name="frmList">
				<input type="hidden" id="Flag" name="Flag" value="SORT">
				<input type="hidden" id="sPage" name="sPage" value="<%=Page%>">
				<input type="hidden" id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
				<input type="hidden" id="ChSort" name="ChSort">
				<input type="hidden" id="ChSortCur" name="ChSortCur">

				<div class="tblArea">
					<table>
						<colgroup>
							<col class="w35" />
							<col class="w50" />
							<col class="w100" />
							<col />
							<col class="w165" />
							<col class="w70" />
							<col class="w70" />
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
								<th>이미지</th>
								<th>URL</th>
								<th>게시기간</th>
								<th>노출순서</th>
								<th>수정</th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="7">등록된 플래시 배너가 없습니다.</td>
							</tr>
							<% Else %>
							<% 		
									' TotalCnt(0), AdmSeq(1), AdmID(2), AdmName(3), AdmLevel(4), AdmEmail(5), VisitDate(6), AdmMobile(7), AdmTel(8)
									vNum = RecordCount - ((Page-1) * listRecord) 
									For i = 0 To arrDataNum 
							%>
							<tr>
								<td><input type="checkbox" name="SeqList" id="SeqList_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
								<td><%=vNum%></td>
								<td>
									<% If Right(arrData(2,i),4) = ".swf" Then %>
									<script type="text/javascript"> getFlash("<%=arrData(2,i)%>", 80, 80); </script>
									<% Else %>
									<img src="<%=arrData(2,i)%>" width="80" height="80" alt="" />
									<% End If %>
								</td>
								<td class="alignL"><a href="<%=arrData(3,i)%>" class="regLink" target="<%=Fn_SetDefault(arrData(3,i),"#","_self","_blank")%>"><%=arrData(3,i)%></a></td>
								<td>
									<%=arrData(5,i)%>
									<% If arrData(5,i) <> "" AND arrData(6,i) <> "" Then %> ~ <% End If %>
									<%=arrData(6,i)%>
								</td>
								<td>
									<a href="javascript:changeSort(<%=arrData(4,i)%>,'UP')">▲</a>
									<a href="javascript:changeSort(<%=arrData(4,i)%>,'DN')">▼</a>
								</td>
								<td>
									<span class="btns btn_in"><a href="javascript:goForm('<%=arrData(1,i)%>','MOD');">수정</a></span>
									<span class="btns btn_in"><a href="javascript:goForm('<%=arrData(1,i)%>','REF')">재등록</a></span>
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
			</div>

			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_g"><a href="javascript:;" id="btnChkAll">전체선택</a></span>
					<span class="btns btn_g"><a href="javascript:;" id="btnDel">삭제</a></span>
					<span class="btns btn_b"><a href="javascript:goForm('0','ADD')">신규등록</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="javascript:;" id="btnPreview">미리보기</a></span>
					<span class="btns btn_g"><a href="javascript:;" id="btnClose">x 닫기</a></span>
				</div>
				
				<div class="paging">
					<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
				</div>
			</div>
		</div>
	</div>
</body>
</html>