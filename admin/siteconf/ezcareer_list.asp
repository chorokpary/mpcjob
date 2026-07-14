<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : ezcareer_list.asp / 경력코드 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2013-07-24 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum

	Dim CONF_TBL, CONF_FLD

	Call Main()
	
	Sub Main()

		listPage 		= 10
		listRecord 	= 40	'목록갯수
		Page 		= FN_Req("Page","1")
		
		CONF_TBL = "TBL_USER_CAREER_EZ"
		CONF_FLD = "Career"

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
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "LIST"
			.Parameters("@Tbl").Value 		= CONF_TBL
			.Parameters("@Field").Value 	= CONF_FLD
			
			Set Rs = .Execute
		End With

		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' CdKey(0), CdVal(1), Sort(2), IsUse(3)
			kFields 	= Array("CdKey","CdVal","Sort","IsUse")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>인력업체관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkIsUse(){
			if (cntChecked("#frmList")==0){
				alert("선택된 데이터가 없습니다.");
			}else{
				if(confirm("사용 여부를 변경하시겠습니까?")){
					$("#Flag").val("STA");
					$("#frmList").attr({action:"ezcareer_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function openForm(sort){
			if (sort=="0" && $("#CodeLen").val() > 30000){
				alert("최대 30000개까지 등록 가능합니다.");
				return;
			}
				
			var url = "ezcareer_form.asp";
			var param = "Sort=" + sort;
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ 
					$("#divInfo").html(content);
					$("#divInfo").show();
					$("#divInfo").focus();
				}, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
		function changeSort(sort,sortcur){
			$("#Flag").val("SORT");
			$("#ChSort").val(sort);
			$("#ChSortCur").val(sortcur);
			$("#frmList").attr({action:"ezcareer_proc.asp", method:"post"}).submit();
		}
		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnStatus").click(chkIsUse);
		    $("#btnReg").click(function(){ openForm('0');});
		    $("#frmList tbody tr").hover(function(){ $(this).css("background-color","#eeeeee"); },function(){ $(this).css("background-color","#ffffff"); })
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
					<h2>경력코드관리</h2>


					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a></span>
						</div>
						<div class="fl_r">
							</span><span class="btns btn_g"><a href="#btnStatus" id="btnStatus">사용 여부 변경</a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">경력 등록</a></span>
						</div>
					</div>
					
					<div class="tblArea">
						<form name="frmList" id="frmList">
						<input type="hidden" id="CodeLen" name="CodeLen" value="<%=arrDataNum+1%>">
						<input type="hidden" id="Flag" name="Flag">
						<input type="hidden" id="ChSort" name="ChSort">
						<input type="hidden" id="ChSortCur" name="ChSortCur">
						<table>
							<colgroup>
								<col class="w40"/>
								<col/>
								<col class="w85"/>
								<col class="w80"/>
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
									<th>경력</th>
									<th>사용 여부</th>
									<th>순서</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="4">등록된 데이터가 없습니다.</td>
								</tr>
								<% Else %>
								<% 
										' CdKey(0), CdVal(1), Sort(2), IsUse(3)
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Sort" id="Sort_<%=arrData(2,i)%>" value="<%=arrData(2,i)%>" title="선택" /></td>
									<td class="alignL"><a href="javascript:openForm('<%=arrData(2,i)%>')"><%=arrData(1,i)%></a></td>
									<td>
										<%=Fn_SetDefault(arrData(3,i),"True","사용","사용안함")%>
									</td>
									<td>
										<a href="javascript:changeSort(<%=arrData(2,i)%>,'UP')">▲</a>
										<a href="javascript:changeSort(<%=arrData(2,i)%>,'DN')">▼</a>
									</td>
								</tr>
								<%
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
						</form>
					</div>
					<p class="note_c">* 코드정보는 다른 데이터와 연동되므로 삭제가 불가능 합니다. 사용여부로 노출될 데이터를 구분해 주세요.</p>

					<div id="divInfo"><!-- 등록, 수정폼 영역 //--></div>
					
				</div>
			</div>
		</div>
		
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>