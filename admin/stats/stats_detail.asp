<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : stats.asp / 회원 통계 - 시도별
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim sBDate, sEDate, Sido
	Dim kFields, i
	Dim arrLocal, arrLocalNum

	Call Main()
	
	Sub Main()

		sBDate 	= FN_Req("sBDate","2006-01-01")
		sEDate 	= FN_Req("sEDate",Date())
		Sido 	= FN_Req("Sido","")
		
		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		
		Call getData()
		
	End Sub
	
	Sub getData()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Statistics"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "MEMSI"
			.Parameters("@BDate").Value 	= sBDate
			.Parameters("@EDate").Value 	= sEDate
			.Parameters("@Sido").Value 	= Sido
			
			Set Rs = .Execute
		End With

		If Rs.Eof or Rs.Bof Then
			arrLocalNum 	= -1
		Else
			' AddrGugun(0), Cnt(1)
			kFields 	= Array("AddrGugun","Cnt")
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
	<title>지역별 상세보기 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function doDownXls(){
			$("#frmPage").attr({action:"stats_detail_xls.asp", method:"post"}).submit();
		}

		$(function() {
			$("#btnXls").click(doDownXls);
		});
	</script>
</head>

<body class="pop stats">
	<div id="wrap">
		<div id="contArea">
			<h2><%=Sido%> 상세 구분</h2>
			<div class="tblArea">
				<table>
					<colgroup>
						<col class="w220" />
						<col />
					</colgroup>
					<tbody>
						<% 		
						' AddrGugun(0), Cnt(1)
						For i = 0 To arrLocalNum 
						%>
						<tr>
							<th bgcolor="#cccccc"><%=arrLocal(0,i)%></th>
							<td class="alignR"><%=arrLocal(1,i)%> 명 </td>
						</tr>
						<% Next %>
					</tbody>
				</table>
			</div>
			<div class="btnArea">
				<div class="fl_r">
					<span class="btns btn_b"><a href="#btnXls" id="btnXls">엑셀저장</a>
					</span><span class="btns btn_g"><a href="#btnClose" id="btnClose">x 닫기</a></span>
				</div>
			</div>
			<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="sBDate" name="sBDate" value="<%=sBDate%>" />
				<input type="hidden" id="sEDate" name="sEDate" value="<%=sEDate%>" />
				<input type="hidden" id="Sido" name="Sido" value="<%=Sido%>" />
			</form>

		</div>
	</div>
</body>
</html>