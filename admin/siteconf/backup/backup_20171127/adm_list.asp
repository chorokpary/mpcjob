<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : adm_list.asp / 관리자 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, sLevel
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim arradminLevel(3)
	
	arradminLevel(1) = "최고관리자"
	arradminLevel(2) = "최고관리자2"
	arradminLevel(3) = "일반관리자"

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 10	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		sLevel 	= FN_Req("sLevel","")
		
		
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
			.CommandText ="uspAdm_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@sLevel").Value 	= sLevel
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), AdmSeq(1), AdmID(2), AdmName(3), AdmLevel(4), AdmEmail(5), VisitDate(6), AdmMobile(7), AdmTel(8)
			kFields 	= Array("TotalCnt","AdmSeq","AdmID","AdmName","AdmLevel","AdmEmail","VisitDate","AdmMobile","AdmTel")
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
	<title>관리자등록 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 관리자가 없습니다.");
			}else{
				if(confirm("선택한 관리자를 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"adm_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function openForm(seq){
			openWin("adm_form.asp?Seq="+seq,"popAdm",0,0,930,400,0)
		}
		
		
		$(function() {
		    $("#sLevel").change(initPage);
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
		    $("#btnReg").click(function(){ openForm(0);});
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
					<h2>관리자등록</h2>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">관리자추가</a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<input type="hidden" id="sKey" name="sKey" value="NAME">
							■ 관리자 구분
							<select id="sLevel" name="sLevel" title="관리자 선택">
								<option value="">선택</option>
								<option value="1"<%=Fn_SetDefault(sLevel,"1"," selected","")%>>최고관리자</option>
								<option value="2"<%=Fn_SetDefault(sLevel,"2"," selected","")%>>최고관리자2</option>
								<option value="3"<%=Fn_SetDefault(sLevel,"3"," selected","")%>>일반관리자</option>
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
								<col class="w35" />
								<col class="w35" />
								<col class="w80" />
								<col class="w65" />
								<col class="w60" />
								<col />
								<col class="w125" />
								<col class="w95" />
								<col class="w95" />
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
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th>관리자 구분</th>
									<th>아이디</th>
									<th>이름</th>
									<th>이메일</th>
									<th>최근접속</th>
									<th>휴대폰</th>
									<th>연락처</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="9">등록된 관리자가 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), AdmSeq(1), AdmID(2), AdmName(3), AdmLevel(4), AdmEmail(5), VisitDate(6), AdmMobile(7), AdmTel(8)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td>
									<!-- <%=Fn_SetDefault(arrData(4,i),"1","최고관리자","일반관리자")%> -->
									<%=arradminLevel(arrData(4,i))%>
									</td>
									<td><a href="javascript:openForm('<%=arrData(1,i)%>')"><%=arrData(2,i)%></a></td>
									<td><a href="javascript:openForm('<%=arrData(1,i)%>')"><%=arrData(3,i)%></a></td>
									<td class="alignL"><%=arrData(5,i)%></td>
									<td><%=arrData(6,i)%></td>
									<td><%=arrData(7,i)%></td>
									<td><%=arrData(8,i)%></td>
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