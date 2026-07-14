<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : stats_detail_xls.asp / 회원 통계 - 시도별 다운로드
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-19 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition","attachment; filename=MPC_STATS_SIDO_EXCEL.xls"
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
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
	<title>MPC_STATS_SIDO_EXCEL</title>
	<style type="text/css">
		th { font-size:9pt; font-weight:bold; text-align:center;}
		td { font-size:9pt; height:20.1pt; display:block; padding:2px; text-align:right;}
	</style>
	</head>
</head>
<body link=blue vlink=purple>

	<h3><%=Sido%> 상세 구분</h3>
	<table border=1 cellpadding=0 cellspacing=0>
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
				<th bgcolor="#cccccc"><%=arrLocal(0,i)%></th>
				<td class="alignR"><%=arrLocal(1,i)%> 명 </td>
			</tr>
			<% Next %>
		</tbody>
	</table>
</div>

</body>
</html>