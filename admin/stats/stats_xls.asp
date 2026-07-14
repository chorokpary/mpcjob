<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : stats_xls.asp / 회원 통계 다운로드
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
	Response.AddHeader "Content-Disposition","attachment; filename=MPC_STATS_EXCEL.xls"
%>
<%
	Dim sBDate, sEDate, IsSearch
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
			arrAgeNum 	= -1
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
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
	<title>MPC_STATS_EXCEL</title>
	<style type="text/css">
		th { font-size:9pt; font-weight:bold; text-align:center; }
		td { font-size:9pt; height:20.1pt; display:block; padding:2px; text-align:right; }
		
		.fl_l {float : left;}
		.fl_r {float : right;}
	</style>
	</head>
</head>
<body link=blue vlink=purple>

	<div class="tblArea">
		<div class="fl_l">
			<h3>성별 통계</h3>
			<table border=1 cellpadding=0 cellspacing=0>
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
						<th bgcolor="#cccccc">
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
			<table border=1 cellpadding=0 cellspacing=0>
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
						<th bgcolor="#cccccc"><%=Fn_SetDefault(arrAge(0,i),"10","기타",arrAge(0,i) & "0 대")%></th>
						<td class="alignR"><%=arrAge(1,i)%> 명</td>
					</tr>
					<% Next %>
				</tbody>
			</table>
		</div>
		
		<div class="fl_r">
			<h3>거주지별 통계</h3>
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
	</div>

</body>
</html>