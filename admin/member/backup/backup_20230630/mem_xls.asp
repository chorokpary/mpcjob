<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_xls.asp / 회원 목록 다운로드
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition","attachment; filename=MPCHR_EXCEL.xls"
%>
<%
	Dim Seq, SortKey, SortDir
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strStyle
	
	Call Main()
	
	Sub Main()

		Seq 		= FN_Req("Seq","")
		SortKey 	= FN_Req("SortXlsKey","")
		SortDir 	= FN_Req("SortXlsDir","")
		
		Seq = Replace(Seq," ","")
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs, Rs_d
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandTimeout = 600
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "XLS"
			.Parameters("@StrSeq").Value 	= Seq
			If SortKey <> "" And SortDir = "DESC" Then
				.Parameters("@SortDesc").Value 	= SortKey
			ElseIf SortKey <> "" And SortDir = "ASC" Then
				.Parameters("@SortAsc").Value 	= SortKey
			End If

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' UsrSeq(0), UsrID(1), UsrName(2), Gender(3), Birthday(4)
			' Addr(5), AddrDetail(6), Mobile(7), Email(8), HopePart(9), RegDate(10), ModDate(11), IsBlack(12)
			' CntMemo(13), Status(14), PartnerDepth1(15), PartnerDepth2(16), PrjName(17), strEzCareer(18), AddrSido(19), AddrGugun(20)
			' Recommend_Num(21), Recommend_Name(22), Recommend_Center(23)
			kFields 	= Array("UsrSeq","UsrID","UsrName","Gender","Birthday","Addr","AddrDetail","Mobile","Email","HopePart","RegDate","ModDate","IsBlack","CntMemo","Status","PartnerDepth1","PartnerDepth2","PrjName","strEzCareer","AddrSido","AddrGugun","Recommend_Num","Recommend_Name","Recommend_Center")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True


			'엑셀 다운로드 로그 입력
			With objCmd
				.ActiveConnection = objDbCon
				.CommandType = 4
				.CommandText ="uspAdm_Down"
				.Parameters.Refresh
				.Parameters("@Flag").Value 			= "LOG"
				.Parameters("@AdmSeq").Value 		= Session("ASeq")
				.Parameters("@Type").Value 			= "M"
				.Parameters("@UserIP").Value 		= Request.ServerVariables("REMOTE_ADDR")
				.Parameters("@Etc1").Value 			= ""
				Set Rs_d = .Execute
			End With

		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
	<title>MPCHR_EXCEL-</title>
	<style type="text/css">
		th { font-size:9pt; font-weight:bold; text-align:center; }
		td { font-size:9pt; height:20.1pt; display:block; padding:2px; }
	</style>
	</head>
</head>
<body link=blue vlink=purple>
<table x:str border=1 cellpadding=0 cellspacing=0 >
	<tr height=26 style="mso-height-source:userset;">
		<th bgcolor="#cccccc">이름</th>
		<th bgcolor="#cccccc">생년월일</th>
		<th bgcolor="#cccccc">연령</th>
		<th bgcolor="#cccccc">성별</th>
		<th bgcolor="#cccccc">핸드폰</th>
		<th bgcolor="#cccccc">최근 지원상태</th>
		<th bgcolor="#cccccc">최근 지원경로</th>
		<th bgcolor="#cccccc">최근 지원일</th>
		<th bgcolor="#cccccc">주소</th>
		<th bgcolor="#cccccc">상세주소</th>
		<th bgcolor="#cccccc">이메일</th>
		<th bgcolor="#cccccc">최근지원 프로젝트</th>
		<th bgcolor="#cccccc">경력</th>
		<th bgcolor="#cccccc">추천인사원번호</th>
		<th bgcolor="#cccccc">추천인성명</th>
		<th bgcolor="#cccccc">지원센터명</th>
	</tr>
	<% 
	If kData Then 
		' UsrSeq(0), UsrID(1), UsrName(2), Gender(3), Birthday(4)
		' Addr(5), AddrDetail(6), Mobile(7), Email(8), HopePart(9), RegDate(10), ModDate(11), IsBlack(12)
		' CntMemo(13), Status(14), PartnerDepth1(15), PartnerDepth2(16), PrjName(17), strEzCareer(18)
		' Recommend_Num(21), Recommend_Name(22), Recommend_Center(23)

		set oSeed = Server.CreateObject( "seed.CBC" )
		returnMsg = oSeed.LoadConfig(KEY_PATH)

		For i = 0 To arrDataNum
			' 블랙리스트 체크
			If arrData(12,i) Then
				strStyle = "style=""background:#FFF000"""
			Else
				strStyle = ""
			End If
	%>
	<tr height=26 style='mso-height-source:userset;'>
		<td align="center" <%=strStyle%>><%= arrData(2,i) %></td>
		<td align="center" <%=strStyle%>><%= arrData(4,i) %></td>
		<td align="center" <%=strStyle%>><% If IsNumeric(Left(arrData(4,i),4)) Then %><%=Year(Date) - Left(arrData(4,i),4) + 1%><% Else %> - <% End If %></td>
		<td align="center" <%=strStyle%>>
			<%
				Select Case arrData(3,i) 
					Case "1" : Response.Write "남" 
					Case "2" : Response.Write "여" 
					Case Else : Response.Write "-" 
				End Select 
			%>
		</td>
		<td align="center" <%=strStyle%>><%= oSeed.Decrypt(arrData(7,i)) %></td>
		<td align="center" <%=strStyle%>><%= arrData(14,i) %></td>
		<td align="center" <%=strStyle%>><%=arrData(15,i)%> 
			<% If Not FN_IsBlank(arrData(16,i)) Then %>(<%=arrData(16,i)%>)<% End If %></td>
		<td align="center" <%=strStyle%>><%=FN_SetDateTimeFormat(arrData(11,i),"YYYY-MM-DD")%></td>
		<td <%=strStyle%>><%= arrData(19,i) %>&nbsp;<%= arrData(20,i) %></td>
		<td <%=strStyle%>><%= arrData(6,i) %></td>
		<td align="center" <%=strStyle%>><%= arrData(8,i) %></td>
		<td align="center" <%=strStyle%>><%= arrData(17,i) %></td>
		<td <%=strStyle%>><%=arrData(18,i)%></td>
		<td align="center"><%= arrData(21,i) %></td>
		<td align="center"><%= arrData(22,i) %></td>
		<td align="center"><%= arrData(23,i) %></td>
	</tr>
	<%
		Next
	Else	
	%>
	<tr height=26 style='mso-height-source:userset;height:20.1pt'>
		<td align="center" colspan="16">선택된 회원 정보가 없습니다.</td>
	</tr>
	<%
	End If
	%>	
</table>

</body>
</html>