<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_apply_stats.asp / 지원자 현황
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-19 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq
	Dim kFields, i, j
	Dim arrData, arrDataNum, arrSum(12)
	Dim tmpCnt

	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","")
		Seq = Replace(Seq," ","")
		
		Call getData()
		
	End Sub
	
	Sub getData()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "STATS_APP"
			.Parameters("@StrSeq").Value 	= Seq

			.Parameters("@AdmLevel").Value 	= Session("ALevel")
			.Parameters("@AdmSeq").Value 	= Session("ASeq")
			
			Set Rs = .Execute
		End With

		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
		Else
			' RecrSeq(0), CareerTxt(1), Title(2), AppBDate(3), AppEDate(4), AdmName(5)
			' CntAll(6), Cnt00(7), Cnt01(8), Cnt02(9), Cnt03(10), Cnt04(11)
			' Cnt05(12), Cnt06(13), Cnt07(14), Cnt08(15), PrjName(16)

			' RecrSeq(0), CareerTxt(1), Title(2), AppBDate(3), AppEDate(4), AdmName(5), PrjName(6)
			' CntAll(7), Cnt00(8), Cnt01(9), Cnt02(10), Cnt03(11), Cnt10(12), Cnt04(13)
			' Cnt05(14), Cnt06(15), Cnt07(16), Cnt09(17), Cnt08(18), Cnt11(19)

			kFields 	= Array("RecrSeq","CareerTxt","Title","AppBDate","AppEDate","AdmName","PrjName","CntAll","Cnt00","Cnt01","Cnt02","Cnt03","Cnt10","Cnt04","Cnt05","Cnt06","Cnt07","Cnt09","Cnt08","Cnt11")
			arrData 		= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>지원자 현황 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function goApply(seq, status){
			openWin("job_apply_user.asp?RecrSeq="+ seq + "&sStatus=" + status,"popApplyUser",0,0,930,620,0);
		}
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap" class="applicant recruitApply">
		<div id="contArea">
			<h2>지원자 현황</h2>
			<div class="inner">
				<% 		
				' 합계 초기화
				For j = 0 To 12
					arrSum(j) = 0
				Next
				' RecrSeq(0), CareerTxt(1), Title(2), AppBDate(3), AppEDate(4), AdmName(5)
				' CntAll(6), Cnt00(7), Cnt01(8), Cnt02(9), Cnt03(10), Cnt04(11)
				' Cnt05(12), Cnt06(13), Cnt07(14), Cnt08(15), PrjName(16)

				' RecrSeq(0), CareerTxt(1), Title(2), AppBDate(3), AppEDate(4), AdmName(5), PrjName(6)
				' CntAll(7), Cnt00(8), Cnt01(9), Cnt02(10), Cnt03(11), Cnt10(12), Cnt04(13)
				' Cnt05(14), Cnt06(15), Cnt07(16), Cnt09(17), Cnt08(18), Cnt11(19)
				For i = 0 To arrDataNum 
					For j = 0 To 12
						tmpCnt = arrData(7+j,i)
						If IsNull(tmpCnt) Or tmpCnt = "" Then 	tmpCnt = 0
						
						arrSum(j) = arrSum(j) + tmpCnt
					Next
				%>
				<h3>◎ [<%=arrData(6,i)%>] <%=arrData(2,i)%></h3>
				<div class="tblArea">
					<table>
						<colgroup>
							<col span="3" class="w65" />
							<col  />
							<col span="13" class="w45" />
						</colgroup>
						<thead class="small">
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
								<th rowspan="2">고유번호</th>
								<th rowspan="2">모집구분</th>
								<th rowspan="2">관리자</th>
								<th>시작일</th>
								<th rowspan="2">모집<br />(지원)</th>
								<th rowspan="2">접수완료</th>
								<th rowspan="2">서류전형<br />합격</th>
								<th rowspan="2">서류전형<br />불합격</th>
								<th rowspan="2">면접<br />불참자</th>
								<th rowspan="2">결과대기</th>
								<th rowspan="2">최종합격</th>
								<th rowspan="2">최종<br />불합격</th>
								<th rowspan="2">교육참석</th>
								<th rowspan="2">교육<br />미참석</th>
								<th rowspan="2">교육이탈</th>
								<th rowspan="2">입사</th>
								<th rowspan="2">채용<br/>취소</th>
							</tr>
							<tr>
								<th>마감일</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td rowspan="2"><%=arrData(0,i)%></td>
								<td rowspan="2"><%=arrData(1,i)%></td>
								<td rowspan="2"><%=arrData(5,i)%></td>
								<td><%=FN_SetDateTimeFormat(arrData(3,i),"YYYY-MM-DD")%></td>
								<% For j = 0 To 12 %>
								<td rowspan="2"><a href="javascript:goApply(<%=arrData(0,i)%>,'<%=Replace(kFields(7+j),"Cnt","")%>')"><%=arrData(7+j,i)%></td>
								<% Next %>
							</tr>
							<tr>
								<td><%=FN_SetDateTimeFormat(arrData(4,i),"YYYY-MM-DD")%></td>
							</tr>
						</tbody>
					</table>
				</div>
				<% Next %>
			</div>
			<div class="tblArea">
				<table>
					<colgroup>
						<col span="4" />
						<col span="13" class="w45" />
					</colgroup>
					<thead>
						<tr class="design">
							<td colspan="4"></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
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
					</thead>
					<tbody>
						<tr>
							<th colspan="4">총 합계</th>
							<% For i = 0 To 12 %>
							<td><%=arrSum(i)%></td>
							<% Next %>
						</tr>
					</tbody>
				</table>
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