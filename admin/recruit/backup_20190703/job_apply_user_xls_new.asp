<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->

<%
'____________________________________________________________________________________
'
' * Discription : job_apply_user_xls_new.asp / 지원자 목록 다운로드 2
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2018-12-21 / 이길홍 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition","attachment; filename=MPC_APPLY_EXCEL.xls"
%>
<%
	Dim Seq, RecrSeq, SortKey, SortDir
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strAppCStyle, strAppCMsg
	
	Call Main()
	
	Sub Main()

		Seq 		= FN_Req("Seq","")
		RecrSeq 	= FN_Req("RecrSeq","")
		SortKey 	= FN_Req("SortXlsKey","")
		SortDir 	= FN_Req("SortXlsDir","")
		
		Seq = Replace(Seq," ","")

		If Session("ALevel") <> "1" And Session("ALevel") <> "2" Then	' 최종관리자가 아닐경우 접근 권한 체크_ 최고관리자2 추가
			Call chkAuth()
		End If
		

		Call getList()
		
	End Sub

	Sub chkAuth()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruitCharge_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "AUTH"
			.Parameters("@Seq").Value 		= RecrSeq
			.Parameters("@AdmSeq").Value 	= Session("ASeq")

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			'권한 없음
			Call SB_ReturnErr("열람 권한이 없는 공고입니다.","CLOSE")
		End If
	
		Rs.Close
		Set Rs = Nothing

	End Sub
		
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "XLS2"
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
			' UsrSeq(0), UsrID(1), UsrName(2), Gender(3), Birthday(4), Addr(5), AddrDetail(6), Mobile(7), Email(8), RegDate(9)
			' CntMemo(10), Status(11), PartnerDepth1(12), PartnerDepth2(13), AddrSido(14), AddrGugun(15), CancelDate(16), IsOut(17), PrjName(18), Times(19)
			kFields 	= Array("UsrSeq","UsrID","UsrName","Gender","Birthday","Addr","AddrDetail","Mobile","Email","RegDate","CntMemo","Status","PartnerDepth1","PartnerDepth2","AddrSido","AddrGugun","CancelDate","IsOut","PrjName","Times")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
	<title>MPC_APPLY_EXCEL</title>
	<style type="text/css">
		th { font-size:9pt; font-weight:bold; text-align:center; }
		td { font-size:9pt; height:20.1pt; display:block; padding:2px; }
		tr.cancel td{background-color:#eeeeee; color:#999999;}
	</style>
	</head>
</head>
<body link=blue vlink=purple>
<table x:str border=1 cellpadding=0 cellspacing=0 >
	<tr height=26 style="mso-height-source:userset;">
		<th bgcolor="#cccccc">이름</th>
		<th bgcolor="#cccccc">의뢰회차</th>
		<th bgcolor="#cccccc">이름</th>
		<th bgcolor="#cccccc">생년월일</th>
		<th bgcolor="#cccccc">연령</th>
		<th bgcolor="#cccccc">성별</th>
		<th bgcolor="#cccccc">핸드폰</th>
		<th bgcolor="#cccccc">지원상태</th>
		<th bgcolor="#cccccc">지원경로</th>
		<th bgcolor="#cccccc">지원일</th>
		<th bgcolor="#cccccc">주소</th>
		<th bgcolor="#cccccc">상세주소</th>
		<th bgcolor="#cccccc">이메일</th>
	</tr>
	<% 
	If kData Then 

			set oSeed = Server.CreateObject( "seed.CBC" )
			returnMsg = oSeed.LoadConfig(KEY_PATH)

		' UsrSeq(0), UsrID(1), UsrName(2), Gender(3), Birthday(4), Addr(5), AddrDetail(6), Mobile(7), Email(8), RegDate(9)
		' CntMemo(10), Status(11), PartnerDepth1(12), PartnerDepth2(13), AddrSido(14), AddrGugun(15), CancelDate(16), IsOut(17)
		
		For i = 0 To arrDataNum
			strAppCStyle = "" : strAppCMsg = ""

			If Not IsNull(arrData(16,i)) Then
				strAppCStyle = " class=""cancel"""
				strAppCMsg = " / 지원취소"
			End If

			If (arrData(17,i)) Then
				strAppCStyle = " class=""cancel"""
				strAppCMsg = " / 탈퇴"
			End If
	%>
	<tr height=26 style='mso-height-source:userset;' <%=strAppCStyle%>>
		<td align="center"><%= arrData(18,i) %></td>
		<td align="center"><%= arrData(19,i) %></td>
		<td align="center"><%= arrData(2,i) %></td>
		<td align="center"><%= arrData(4,i) %></td>
		<td align="center"><% If IsNumeric(Left(arrData(4,i),4)) Then %><%=Year(Date) - Left(arrData(4,i),4) + 1%><% Else %> - <% End If %></td>
		<td align="center">
			<%
				Select Case arrData(3,i) 
					Case "1" : Response.Write "남" 
					Case "2" : Response.Write "여" 
					Case Else : Response.Write "-" 
				End Select 
			%>
		</td>
		<td align="center"><%= oSeed.Decrypt(arrData(7,i)) %></td>
		<td align="center"><%= arrData(11,i) %><%=strAppCMsg%></td>
		<td align="center"><%=arrData(12,i)%> 
			<% If Not FN_IsBlank(arrData(13,i)) Then %>(<%=arrData(13,i)%>)<% End If %></td>
		<td align="center"><%=FN_SetDateTimeFormat(arrData(9,i),"YYYY-MM-DD")%></td>
		
		<% If FN_isBlank(arrData(5,i)) Then %>
		<td><%= arrData(14,i) %>&nbsp;<%= arrData(15,i) %></td>
		<% Else %>
		<td><%= arrData(5,i) %></td>
		<% End If %>
		
		<td><%= arrData(6,i) %></td>
		<td align="center"><%= arrData(8,i) %></td>
	</tr>
	<%
		Next
	Else	
	%>
	<tr height=26 style='mso-height-source:userset;height:20.1pt'>
		<td align="center" colspan="13">선택된 정보가 없습니다.</td>
	</tr>
	<%
	End If
	%>	
</table>

</body>
</html>