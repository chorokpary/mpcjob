<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : stats.asp / 회원 통계
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim sBDate, sEDate
	Dim kFields, i
	Dim arrGender, arrGenderNum
	Dim arrAge, arrAgeNum
	Dim arrLocal, arrLocalNum

	Call Main()
	
	Sub Main()

		sBDate 	= FN_Req("sBDate","2006-01-01")
		sEDate 	= FN_Req("sEDate",Date())
		
		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		
		If IsDate(sBDate) And IsDate(sEDate) Then
			Call getData()
		Else
			arrGenderNum = -1
			arrAgeNum = -1
			arrLocalNum = -1
		End If
		
	End Sub
	
	Sub getData()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Statistics"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "MEM"
			.Parameters("@BDate").Value 	= sBDate
			.Parameters("@EDate").Value 	= sEDate
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrGenderNum 	= -1
		Else
			' Gender(0), Cnt(1)
			kFields 	= Array("Gender","Cnt")
			arrGender 		= Rs.GetRows(,,kFields)
			arrGenderNum 	= UBound(arrGender, 2)
		End If
		
		Set Rs = Rs.NextRecordSet()

		If Rs.Eof or Rs.Bof Then
			arrAgeNum 	= -1
		Else
			' Age(0), Cnt(1)
			kFields 	= Array("Age","Cnt")
			arrAge 		= Rs.GetRows(,,kFields)
			arrAgeNum 	= UBound(arrAge, 2)
		End If

		Set Rs = Rs.NextRecordSet()

		If Rs.Eof or Rs.Bof Then
			arrLocalNum 	= -1
		Else
			' SIDO(0), Cnt(1)
			kFields 	= Array("SIDO","Cnt")
			arrLocal 		= Rs.GetRows(,,kFields)
			arrLocalNum 	= UBound(arrLocal, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>HC 회원통계 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript">
		function chkSearch() {
			if (!$("#sBDate").val()){
				alert("검색 범위를 입력해주세요.");
				$("#sBDate").focus();
				return;
			} else if (!$("#sEDate").val()){
				alert("검색 범위를 입력해주세요.");
				$("#sEDate").focus();
				return;
			}
			$("#frmPage").attr({action:"stats.asp", method:"post", target:"_self"}).submit();
		}
		
		function doDownXls(){
			$("#frmPage").attr({action:"stats_xls.asp", method:"post"}).submit();
		}
		
		function goDetail(Sido){
			openWin("about:blank","popDetail",0,0,940,650,1);
			$("#Sido").val(Sido);
			$("#frmPage").attr({action:"stats_detail.asp", method:"post", target:"popDetail"}).submit();
		}

		$(function() {
		    $("#btnSearch").click(chkSearch);
		    $("#btnXls").click(doDownXls);
		    
			$("#sBDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#sEDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});

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
				
				<div id="contArea" class="stats">
					<h2>HC 회원통계</h2>
					<div class="searchArea">
						<form name="frmPage" id="frmPage" method="post">
						<div class="fl_l">
							<span class="tit_date">기간설정</span> 
							<input type="text" id="sBDate" name="sBDate" value="<%=sBDate%>" class="w70" style="margin-right:5px;" maxlength="10" /> 부터 ~ 
							<input type="text" id="sEDate" name="sEDate" value="<%=sEDate%>" class="w70" style="margin-right:5px;" maxlength="10" /> 까지
							<input type="hidden" id="Sido" name="Sido">
							<span class="btns btn_b"><a href="#btnSearch" id="btnSearch">검색</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#btnXls" id="btnXls">엑셀저장</a></span>
						</div>
						</form>
					</div>
					<p class="note_c">* 입력되지 않은 정보는 통계에 포함되지 않습니다.</p>
					<!-- p class="note_c">* 30일 이상은 검색 하실수 없습니다.</p //-->
					
					<div class="tblArea">
						<div class="fl_l">
							<h3>성별 통계</h3>
							<table>
								<colgroup>
									<col class="w100" />
									<col />
								</colgroup>
								<tbody>
									<% 		
									' Gender(0), Cnt(1)
									For i = 0 To arrGenderNum 
									%>
									<tr>
										<th>
										<%
											If arrGender(0,i) = "1" Then
												Response.Write "남성"
											ElseIf arrGender(0,i) = "2" Then
												Response.Write "여성"
											End If
										%>
										</th>
										<td class="alignR"><%=arrGender(1,i)%> 명</td>
									</tr>
									<% Next %>
								</tbody>
							</table>
							<h3>연령 통계</h3>
							<table>
								<colgroup>
									<col class="w100" />
									<col />
								</colgroup>
								<tbody>
									<% 		
									' Age(0), Cnt(1)
									For i = 0 To arrAgeNum 
									%>
									<tr>
										<th><%=Fn_SetDefault(arrAge(0,i),"10","기타",arrAge(0,i) & "0 대")%></th>
										<td class="alignR"><%=arrAge(1,i)%> 명</td>
									</tr>
									<% Next %>
								</tbody>
							</table>
						</div>
						
						<div class="fl_r">
							<h3>거주지별 통계</h3>
							<table>
								<colgroup>
									<col class="w100" />
									<col />
								</colgroup>
								<tbody>
									<% 		
									' SIDO(0), Cnt(1)
									For i = 0 To arrLocalNum 
									%>
									<tr>
										<th><%=arrLocal(0,i)%></th>
										<td class="alignR"><%=arrLocal(1,i)%> 명 <span class="btns btn_in"><a href="javascript:goDetail('<%=arrLocal(0,i)%>')">상세보기</a></span></td>
									</tr>
									<% Next %>
								</tbody>
							</table>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>