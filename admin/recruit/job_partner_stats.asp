<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_patner_stats.asp / 통계관리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-22 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, ViewFlag
	Dim kFields, i, j
	Dim arrData, arrDataNum
	Dim arrList, arrListNum


	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","")
		ViewFlag = FN_Req("ViewFlag","SING")
		
		Seq = Replace(Seq," ","")
		
	End Sub

	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "MUL_SEL"
			.Parameters("@StrSeq").Value 	= Seq

			.Parameters("@AdmSeq").Value 	= Session("ASeq")
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrListNum 	= -1
		Else
			' RecrSeq(0), Title(1), AdmSeq(2), AppBDate(3), AppEDate(4), PrjName(5)

			kFields 	= Array("RecrSeq","Title","AdmSeq","AppBDate","AppEDate","PrjName")
			arrList 	= Rs.GetRows(,,kFields)
			arrListNum 	= UBound(arrList, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing
		
		' 컨텐츠 출력
		For j = 0 To arrListNum
			If Session("ALevel") = "1" Or (Session("ALevel") = "3" And Session("ASeq") = arrList(2,j)) Then
				'Response.Write "<h3>◎ " & arrList(1,j) & "</h3>"
				Response.Write "<h3>◎ " & arrList(5,j) &  " (" & LEFT(arrList(3,j),10) & " ~ " & LEFT(arrList(4,j),10) &  ")</h3>"
				Call getData(arrList(0,j))
			End If
		Next

	End Sub
		
	Sub getData(argSeq)

		Dim Rs, PreSeq
		Dim CntAll, Cnt1, Cnt2, Cnt3, Cnt4, Cnt5
		Dim Per1, Per2, Per3, Per4
		Dim StrStyle
		
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
	<title>통계관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function doDownXls(){
			$("#frmInfo").attr({action:"job_partner_stats_xls.asp", method:"post", target:"_self"}).submit();
		}
		
		$(function() {
		    $("#btnXls").click(doDownXls);
		});
	</script>
</head>

<body class="pop">
	<div id="wrap" class="statistics">
		<div id="contArea">
			<h2>통계관리</h2>
			<div class="tabArea">
				<span class="tab<%=Fn_SetDefault(ViewFlag,"SING"," on","")%>"><a href="job_partner_stats.asp?ViewFlag=SING&Seq=<%=Seq%>">개별통계</a>
				</span><span class="tab<%=Fn_SetDefault(ViewFlag,"SUM"," on","")%>"><a href="job_partner_stats.asp?ViewFlag=SUM&Seq=<%=Seq%>">전체통계</a></span>
			</div>
			
			<div class="inner">
				<!-- 반복 : start //-->
				<%
				If ViewFlag = "SING" Then
					Call getList()
				Else
					Call getData(Seq)
				End If
				%>
				<!-- 반복 : end //-->
				<form id="frmInfo" name="frmInfo">
					<input type="hidden" id="ViewFlag" name="ViewFlag" value="<%=ViewFlag%>">
					<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
				</form>
			</div>
			<div class="btnArea">
				<div class="fl_r">
					<span class="btns btn_b"><a href="#btnXls" id="btnXls">EXCEL 변환</a>
					</span><span class="btns btn_g"><a href="#btnClose" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>