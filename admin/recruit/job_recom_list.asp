<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_recom_list.asp / 추천공고 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-22 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim sViewFlag
	Dim StrViewFlag : StrViewFlag = ""
	Dim StrTitle : StrTitle = ""

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 24	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))

		sViewFlag = FN_Req("sViewFlag","TOP")

		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		If sViewFlag = "TOP" Then
			listRecord = 4
			StrTitle = "스페셜 HOT"
		Else
			listRecord = 24
			StrTitle = "이달의 추천공고"
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 20
		End If
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord

			.Parameters("@sIsView").Value 	= 1
			.Parameters("@sViewFlag").Value = sViewFlag

			.Parameters("@AdmLevel").Value 	= Session("ALevel")
			.Parameters("@AdmSeq").Value 	= Session("ASeq")
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), RecrSeq(1), Title(2), IsMain(3), IsTop(4), AppBDate(5), AppEDate(6)
			' PrjName(7), AdmName(8), AppCnt(9)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","IsMain","IsTop","AppBDate","AppEDate","PrjName","AdmName","AppCnt")
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
	<title>추천공고 관리 - MPC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function goApply(seq){
			openWin("job_apply_user.asp?RecrSeq="+ seq,"popApplyUser",0,0,930,620,0);
		}

		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 추천 공고가 없습니다.");
			}else{
				if(confirm("선택한 추천 공고를 삭제하시겠습니까?")){
					$("#Flag").val("NO_RECOM");
					$("#frmList").attr({action:"job_proc.asp", method:"post"}).submit();
				}
			}
		}
		

		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
		});
	</script>
</head>

<body class="pop">
	<div id="wrap" class="monthNotice">
		<div id="contArea">
			<div class="tabArea">
				<span class="tab<%=Fn_SetDefault(sViewFlag,"TOP"," on","")%>"><a href="job_recom_list.asp?sViewFlag=TOP">스페셜 HOT</a>
				</span><span class="tab<%=Fn_SetDefault(sViewFlag,"MAIN"," on","")%>"><a href="job_recom_list.asp?sViewFlag=MAIN">이달의 추천공고</a></span>
			</div>
			<h2><%=StrTitle%> 관리 (<%=arrDataNum+1%>/<%=listRecord%>)</h2>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
					</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a></span>
				</div>
			</div>
			<div class="inner">
				<form name="frmList" id="frmList">
				<input type="hidden" id="Flag" name="Flag">
				<input type="hidden" id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
				<div class="tblArea">
					<table>
						<colgroup>
							<col class="w40" />
							<col class="w40" />
							<col class="w80" />
							<col />
							<col class="w85" />
							<col class="w60" />
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
								<th>담당자</th>
								<th>프로젝트명</th>
								<th>노출</th>
								<th>총 지원자</th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="6">등록된 추천 공고가 없습니다.</td>
							</tr>
							<% Else %>
							<% 		
									' TotalCnt(0), RecrSeq(1), Title(2), IsMain(3), IsTop(4), AppBDate(5), AppEDate(6)
									' PrjName(7), AdmName(8), AppCnt(9)
									vNum = RecordCount - ((Page-1) * listRecord) 
									
									For i = 0 To arrDataNum 
										If arrData(3,i) Or arrData(4,i) Then
											If arrData(3,i) Then
												StrViewFlag = "이달의 추천공고"
											End If
											StrViewFlag = Fn_SetDefault(arrData(3,i),True,"이달의 추천공고","")
											
											If arrData(4,i) Then
												If StrViewFlag <> "" Then 		StrViewFlag = StrViewFlag & "<br>"
												StrViewFlag = StrViewFlag & "스페셜 HOT"
											End If
										Else
											StrViewFlag = "일반채용공고"
										End If
							%>
							<tr>
								<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
								<td><%=vNum%></td>
								<td><%=arrData(8,i)%></td>
								<td class="alignL">
									<div class="job_info">
										<strong><%=arrData(7,i)%></strong> <%=arrData(2,i)%>
									</div>
									<div class="job_detail">
										<p>&middot; 전형시작일 : <%=FN_SetDateTimeFormat(arrData(5,i),"YYYY-MM-DD")%></p>
										<p>&middot; 전형마감일 : <%=FN_SetDateTimeFormat(arrData(6,i),"YYYY-MM-DD")%></p>
									</div>
								</td>
								<td><%=StrViewFlag %></td>
								<td><%=arrData(9,i)%><br /><span class="btns btn_in"><a href="javascript:goApply(<%=arrData(1,i)%>);">보기</a></span>
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
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>