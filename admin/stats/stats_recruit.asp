<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : stats_recuit.asp / 채용 통계
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2025-04-02 / 이길홍 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim sBDate, sEDate
	Dim kFields, i
	Dim arrData, arrDataNum

	Dim sWord
	Dim Seq
	Dim ViewFlag 


	Call Main()
	
	Sub Main()
		
		ViewFlag = "SUM"

		sBDate = FN_Req("sBDate", Year(Date()) & "-01-01")
		sEDate 	= FN_Req("sEDate",Date())

		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		
	End Sub

	Sub getSeq()

		Dim Sql, Rss

		Sql = "SELECT STUFF(("
		Sql = Sql & " SELECT ',' + CAST(D.RecrSeq AS NVARCHAR)"
		Sql = Sql & " FROM TBL_RECRUIT AS D WITH (nolock)"
		Sql = Sql & " LEFT OUTER JOIN TBL_PROJECT AS P WITH (nolock)"
		Sql = Sql & " ON D.PrjSeq = P.PrjSeq"
		Sql = Sql & " WHERE CONVERT(nvarchar(10),D.AppBDate,21) >= '"&Replace(sBDate, "-", "")&"' AND CONVERT(nvarchar(10),D.AppBDate,21) <= '"&Replace(sEDate, "-", "")&"'"
		If Trim(sWord) <> "" Then
			Sql = Sql & " AND P.PrjName like '%"&sWord&"%'"
		End If
		Sql = Sql & " ORDER BY D.RecrSeq DESC"
		Sql = Sql & " FOR XML PATH('')), 1, 1, '') AS RecrSeqList;"

		Response.Write "<script>console.log('Sql : " & Sql & "');</script>"

		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 1 ' adCmdText
			.CommandText =Sql
			Set Rss = .Execute			
		End With

		 If Rss.eof = false  Then    '레코드셋에 값이 있으면 계속 반복
			Seq = Rss(0)			
		 Else
			Seq = Rss(0)
		 End If 

		Call getData(Seq)
	End Sub 

	Sub getData(argSeq)

		Dim Rs, PreSeq
		Dim CntAll, Cnt1, Cnt2, Cnt3, Cnt4, Cnt5
		Dim Per1, Per2, Per3, Per4
		Dim StrStyle

		
		Response.Write "<script>console.log('argSeq : " & argSeq & "');</script>"
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "STATS_ROOT"
			.Parameters("@StrSeq").Value 	= argSeq

			Set Rs = .Execute
		End With

		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			Response.Write "<script>console.log('Rs Empty!!');</script>"
		Else
			' P1_Seq(0), P2_Seq(1), P1_ParnterName(2), P2_ParnterName(3), P1_IsSum(4), P2_IsSum(5), CntAll(6), CntSub(7)
			' Cnt00(8), Cnt01(9), Cnt02(10), Cnt03(11), Cnt04(12), Cnt05(13), Cnt06(14), Cnt07(15), Cnt08(16), Cnt09(17)

			kFields 	= Array("P1_Seq","P2_Seq","P1_ParnterName","P2_ParnterName","P1_IsSum","P2_IsSum","CntAll","CntSub","Cnt00","Cnt01","Cnt02","Cnt03","Cnt04","Cnt05","Cnt06","Cnt07","Cnt08","Cnt09")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

		' 컨텐츠 출력
		%>
				<div class="tblArea">
					<table>
						<colgroup>
							<col span="10" />
						</colgroup>
						<tbody>
							<tr>
								<th colspan="2" rowspan="2">채용루트</th>
								<th>지원(접수)자</th>
								<th>서류합격자</th>
								<th colspan="2">면접참석자</th>
								<th colspan="2">합격자</th>
								<th colspan="2">교육참석</th>
								<th colspan="2">입사</th>
							</tr>
							<tr>
								<th>인원</th>
								<th>인원</th>
								<th>인원</th>
								<th>참석률(%)</th>
								<th>인원</th>
								<th>합격률(%)</th>
								<th>인원</th>
								<th>참석률(%)</th>
								<th>인원</th>
								<th>입사율(%)</th>
							</tr>
							<%
							'면접 참석률 : 면접참석자/서류합격자
							'합격률 : 합격자/면접률 -> 최종합격자 / 서류합격자
							'교육 참석율 : 교육참석자 / 서류합격자
							'입사율 : 채용인원/합격률 -> 입사 / 서류합격자
							PreSeq = 0

							For i = 0 To arrDataNum
								If IsNull(arrData(7,i)) And arrData(5,i) = "1" And arrData(4,i) = "0" Then
								' 2Depth가 없는데 소계가 계산된 내용 제외
								Else
									CntAll = arrData(6,i)	'지원자
									Cnt5 = arrData(16,i)	'채용인원 (입사건)
									Cnt4 = arrData(14,i) + arrData(17,i) + Cnt5	'교육참석인원 (교육참석+교육이탈+[채용인원])
									Cnt3 = arrData(12,i) + arrData(15,i) + Cnt4	'합격자 (최종합격 + 교육미참석 + [교육참석인원])
									Cnt2 = arrData(13,i) + Cnt3 	'면접 참석자 (최종불합격 + [합격자]) 
									Cnt1 = arrData(9,i) + arrData(11,i) + Cnt2 	'서류합격자 (서류전형합격 + 면접불참석 + [면접참석자]) 
	
									'면접참석률 (면접참석자/서류합격자)
									If Cnt1 > 0 And Cnt2 > 0 Then 
										Per1 = "(" & FormatPercent(Cnt2 / Cnt1, 2) & ")"
									Else
										Per1 = ""
									End If
									
									' 합격률 (최종합격자 / 서류합격자)
									If Cnt1 > 0 And Cnt3 > 0 Then 
										Per2 = "(" & FormatPercent(Cnt3 / Cnt1, 2) & ")"
									Else
										Per2 = ""
									End If

									' 교육 참석율 (교육참석자 / 서류합격자)
									If Cnt1 > 0 And Cnt4 > 0 Then 
										Per3 = "(" & FormatPercent(Cnt4 / Cnt1, 2) & ")"
									Else
										Per3 = ""
									End If
										
									' 입사율 (입사 / 서류합격자)
									If Cnt1 > 0 And Cnt5 > 0 Then 
										Per4 = "(" & FormatPercent(Cnt5 / Cnt1, 2) & ")"
									Else
										Per4 = ""
									End If
									
									If  arrData(4,i)=1 Then
										StrStyle = "class=""mainTotal"""
									ElseIf arrData(4,i)=0 And arrData(5,i)=1 Then
										StrStyle = "class=""subTotal"""
									Else
										StrStyle = ""
									End If
								%>
								<tr <%=StrStyle%>>
									<%
									If IsNull(arrData(7,i)) Then
									' 2Depth가 없는 경우
									%>
									<th colspan="2"><%=arrData(2,i)%></th>
									<%
									Else
									' 2Depth가 있는 경우
										If PreSeq <> arrData(0,i) AND Not IsNull(arrData(0,i)) Then 
									%>
									<th rowspan="<%=arrData(7,i) +1 %>"><%=arrData(2,i)%></th>
									<% 
										End If 
									%>
									<th><%=arrData(3,i)%></th>
									<%
									End If
									%>
									<td><%=Fn_SetDefault(CntAll,"0","",CntAll & "명")%></td>
									<td><%=Fn_SetDefault(Cnt1,"0","",Cnt1 & "명")%></td>
									<td><%=Fn_SetDefault(Cnt2,"0","",Cnt2 & "명")%></td>
									<td><%=Per1%></td>
									<td><%=Fn_SetDefault(Cnt3,"0","",Cnt3 & "명")%></td>
									<td><%=Per2%></td>
									<td><%=Fn_SetDefault(Cnt4,"0","",Cnt4 & "명")%></td>
									<td><%=Per3%></td>
									<td><%=Fn_SetDefault(Cnt5,"0","",Cnt5 & "명")%></td>
									<td><%=Per4%></td>
								</tr>
								<%
							End If
								If Not IsNull(arrData(0,i)) Then 		PreSeq = arrData(0,i)
							Next
							%>
						</tbody>
					</table>
				</div>
		<%
	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>HC 채용통계 - HC 관리자</title>
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
			$("#frmPage").attr({action:"stats_recruit.asp", method:"post", target:"_self"}).submit();
		}
		
		function doDownXls(){
			$("#frmPage").attr({action:"stats_recruit_xls.asp", method:"post"}).submit();
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
					<h2>HC 채용통계</h2>
					<div class="searchArea">
						<form name="frmPage" id="frmPage" method="post">						
						<div class="fl_l">
							<span class="tit_date">기간설정</span> 
							<input type="text" id="sBDate" name="sBDate" value="<%=sBDate%>" class="w70" style="margin-right:5px;" maxlength="10" /> 부터 ~ 
							<input type="text" id="sEDate" name="sEDate" value="<%=sEDate%>" class="w70" style="margin-right:5px;" maxlength="10" /> 까지
							<span class="tit_date">프로젝트명</span> 
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>" title="검색어" class="w190" />
							<span class="btns btn_b"><a href="#btnSearch" id="btnSearch">검색</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#btnXls" id="btnXls">엑셀저장</a></span>
						</div>
						</form>
					</div>
					<p class="note_c">* 입력되지 않은 정보는 통계에 포함되지 않습니다.</p>
					<!-- p class="note_c">* 30일 이상은 검색 하실수 없습니다.</p //-->
					<%
						If IsDate(sBDate) And IsDate(sEDate) Then
							Call getSeq()
						End If 
					%>					
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>