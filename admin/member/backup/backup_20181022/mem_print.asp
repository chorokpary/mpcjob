<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_form.asp / 회원정도 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2014-01-07 / 강미현 / 4M / 희망분야 등록항목 삭제
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page, IsUser, RecrSeq
	Dim UsrName, Birthday, Gender, Mobile, Email, HopePart, HopePart2, UsrType
	Dim Zipcode1, Zipcode2, AddrSido, AddrGugun, Addr, AddrDetail, Jumin1, Jumin2
	Dim ImgPath, Attach1_Path, Attach2_Path, Attach3_Path, Attach1_FName, Attach2_FName, Attach3_FName
	Dim RegDate, ModDate, Intro, EduLast, EduType
	Dim kFields, arrLen, i
	Dim arrEdu, arrCareer, arrFam, arrMemo, arrApply
	Dim arrQues
	
	Dim iBegin, iEnd, iRow
	
	' 학력정보
	Dim strEduDtB, strEduDtE, strEduNm, strEduMajor, strEduStatus, strEduLocal
	' 경력정보
	Dim strCarDt, strCarCNm, strCarPos, strCarDpt, strCarTask, strCarRs	
	' 가족사항
	Dim strFamRel, strFamNm, strFamAge, strFamEdu, strFamJob, strFamWith
	
	' 이력서 정보 Get
	Sub getData(argSeq)
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= argSeq
			
			' ## 회원정보
			Set Rs = .Execute
			
			IsUser = False
			If Not Rs.Eof Then
				IsUser 		= True
				
				UsrName 	= FN_InputVal(Rs("UsrName"))
				Birthday 	= Rs("Birthday")
				Gender 	= Rs("Gender")
				Mobile 	= FN_InputVal(Rs("Mobile"))
				Email 	= FN_InputVal(Rs("Email"))
				'HopePart 	= Rs("HopePart")
				'HopePart2 = Rs("HopePart2")
				UsrType 	= Rs("UsrType")
				Zipcode1 = Rs("Zipcode1")
				Zipcode2 = Rs("Zipcode2")
				AddrSido 	= FN_InputVal(Rs("AddrSido"))
				AddrGugun 	= FN_InputVal(Rs("AddrGugun"))
				Addr 	= FN_InputVal(Rs("Addr"))
				AddrDetail = FN_InputVal(Rs("AddrDetail"))
				Jumin1 	= FN_InputVal(Rs("Jumin1"))
				Jumin2 	= FN_InputVal(Rs("Jumin2"))

				ImgPath 	= Rs("ImgPath")
				Attach1_Path 	= Rs("Attach1_Path")
				Attach1_FName = Rs("Attach1_FName")
				Attach2_Path 	= Rs("Attach2_Path")
				Attach2_FName = Rs("Attach2_FName")
				Attach3_Path 	= Rs("Attach3_Path")
				Attach3_FName = Rs("Attach3_FName")
				
				RegDate = Rs("RegDate")
				ModDate = Rs("ModDate")
				
				Intro = Rs("Intro")

				EduLast = Rs("EduLast")
				EduType = Rs("EduType")
			End If

			Set Rs = Rs.NextRecordSet()
				
			'arrDataNum 	= UBound(arrData, 2)
			If Not Rs.Eof Then
				' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
				kFields 	= Array("UsrSeq","EduBYear","EduBMonth","EduEYear","EduEMonth","SchoolName","Major","Status","Local","Sort")
				arrEdu 	= Rs.GetRows(,,kFields)
			Else
				arrEdu = null
			End If
			
			Set Rs = Rs.NextRecordSet()
				
			If Not Rs.Eof Then
				' UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10), CompanyYear(11)
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort" ,"CompanyYear")
				arrCareer 	= Rs.GetRows(,,kFields)
			Else
				arrCareer = null
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
				kFields 	= Array("UsrSeq","Relation","FmName","Age","Education","Job","IsLiveWith","Sort")
				arrFam 	= Rs.GetRows(,,kFields)
			Else
				arrFam = null
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrMSeq(0), AdmName(1), Memo(2)
				kFields 	= Array("UsrMSeq","AdmName","Memo")
				arrMemo 	= Rs.GetRows(,,kFields)
			Else
				arrMemo = null
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' PrjName(0), Title(1), AppBDate(2), AppEDate(3), RegDate(4), AdmName(5), CdVal(6), RecrPart(7), RecrPart2(8)
				kFields 	= Array("PrjName","Title","AppBDate","AppEDate","RegDate","AdmName","CdVal","RecrPart","RecrPart2")
				arrApply 	= Rs.GetRows(,,kFields)
				
				HopePart = arrApply(7,0)
				HopePart2 = arrApply(8,0)
			Else
				arrApply = null
			End If			

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
	' 사전질문 답변 Get
	Sub getDataQuestion(argSeq)
		Dim Rs, kFields
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = "uspRecruitQaQ_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "LST_FRONT"
			.Parameters("@Seq").Value 	= RecrSeq
			.Parameters("@UsrSeq").Value = argSeq
			
			' ## 회원정보
			Set Rs = .Execute
	
			If Not Rs.Eof Then
				kFields 	= Array("QSeq","QCode","Sort","Question","IsDisp","Answer")
				arrQues = Rs.GetRows(,,kFields)
			Else
				arrQues = null
			End If
	
			Rs.Close
			Set Rs = Nothing
		End With	
	End Sub

	Sub Main()
		Dim arrSeq, k

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		RecrSeq = FN_Req("RecrSeq","0")
		
		If Not IsNumeric(RecrSeq) Then 	RecrSeq = 0
		
		arrSeq = Split(Replace(Seq," ",""),",")
		
		For k = 0 To UBound(arrSeq)
			Call getData(arrSeq(k))
			
			If RecrSeq > 0 Then
				Call getDataQuestion(arrSeq(k))
			End If
			
			If IsUser Then
	
%>
		<div id="contArea">
			<div class="<%=Fn_SetDefault(k,0,"","printSec")%> pb50">
				<h2>개인정보</h2>
				<div class="btnArea mt-20">
					<div class="fl_r">
						<p class="prtName"><%=UsrName%></p><span class="btns btn_g"><a href="#btnPrint" class="btnPrint">인쇄하기</a></span>
					</div>
				</div>
				<div >
					<div class="peopleInfo">
						<div class="fl_l">
							<div class="imgs">
								<div class="inner">
									<% If Not FN_isBlank(ImgPath) Then %>
									<img src="<%=ImgPath%>" width="116" height="144">
									<% Else %>
									<img alt="" src="/admin/images/img_nopics.jpg" />
									<% End If %>
								</div>
							</div>
						</div>
						<div class="fl_r">
							<div class="tblArea">
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col />
										<col class="w85" />
										<col />
									</colgroup>
									<tbody>
										<tr class="design">
											<td></td>
											<td></td>
											<td></td>
											<td></td>
										</tr>
										<tr>
											<th><label for="">이름</label></th>
											<td><%=UsrName%></td>
											<th><label for="">생년월일</label></th>
											<td><%=BirthDay%></td>
										</tr>
										<tr>
											<!-- <th><label for="">주민번호</label></th>
											<td><%=Jumin1%>-<%=Jumin2%></td> -->
											<th><label for="">성별</label></th>
											<td colspan="3">
												<% 
												If Gender="1" Then 
													Response.Write "남자"
												ElseIf Gender = "2" Then
													Response.Write "여자"
												End If
												%>
											</td>
										</tr>
										<tr>
											<th><label for="">아이디</label></th>
											<td><%=Mobile%></td>
											<th><label for="">희망분야</label></th>
											<td><%=HopePart%> <% If HopePart2 <> "" AND Not ISNULL(HopePart2) Then Response.Write " / " & HopePart2 %></td>
										</tr>
										<tr>
											<th><label for="">E-mail 주소</label></th>
											<td><%=Email%></td>
											<th><label for="">핸드폰번호</label></th>
											<td><%=Mobile%></td>
										</tr>
										<tr>
											<th><label for="">주소</label></th>
											<td colspan="3"><%=Fn_SetDefault(Addr,"",AddrSido & " " & AddrGugun,Addr & " " & AddrDetail)%></td>
										</tr>
									</tbody>
								</table>
							</div>
						</div>
					</div>
					<div class="section">
						<h2>학력사항</h2>
						<div class="tblArea">
							<%
							Dim strFinEdu : strFinEdu = ""
							If IsArray(arrEdu) Then
								arrLen = UBound(arrEdu, 2)
								strFinEdu = arrEdu(5,arrLen)
							Else
								arrLen = -1
							End If
							%>
							<table class="tblForm">
								<colgroup>
									<col class="w85" />
									<!-- <col class="w30" />-->
									<col class="w110" /> 
									<col class="w110" />
									<col />
									<!-- <col class="w100" />
									<col class="w80" />
									<col class="w100" /> -->
								</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<!-- <td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td> -->
									</tr>
									<tr>
											<th rowspan="<%=(4+arrLen)%>">최종학력사항</th>
											<td class="alignC">
												<%
												Select Case(EduLast)
													Case "0"
														Response.Write "중학교"
													Case "1"
														Response.Write "고등학교"
													Case "2"
														Response.Write "대학(2,3년)"
													Case "3"
														Response.Write "대학교(4년)"
													Case "4"
														Response.Write "대학원"
												End Select 
												%>
											</td>
											<td class="alignC">
												<%
												Select Case(EduType)
													Case "1"
														Response.Write "졸업"
													Case "2"
														Response.Write "재학중"
													Case "3"
														Response.Write "휴학"
													Case "4"
														Response.Write "중퇴"
												End Select 
												%>
											</td>
											<!-- <th colspan="7" class="alignL"><% If arrLen > -1 Then %>최종학력 : <%=strFinEdu%> <% End If %></th> -->
									</tr>
									<!-- <tr>
										<th>입학년월</th>
										<th>졸업년월</th>
										<th>학교명</th>
										<th>전공</th>
										<th>졸업여부</th>
										<th>소재지</th>
									</tr> -->
									<% 		
									' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
									iRow = 2
									If arrLen <= iRow Then
										iBegin = 0
										iEnd = iRow
									Else
										iBegin = arrLen - iRow
										iEnd = arrLen
									End If
									
									For i = iBegin To iEnd
										strEduDtB = "" : strEduDtE = "" : strEduNm = "" : strEduMajor = "" : strEduStatus = "" : strEduLocal = ""
										
										If i <= arrLen Then
											If arrEdu(1,i) <> "" And arrEdu(2,i) <> "" Then
												strEduDtB = arrEdu(1,i) & " 년 " & arrEdu(2,i) & " 월"
											End If
											If arrEdu(3,i) <> "" And arrEdu(4,i) <> "" Then
												strEduDtE = arrEdu(3,i) & " 년 " & arrEdu(4,i) & " 월"
											End If
											strEduNm = arrEdu(5,i)
											strEduMajor = arrEdu(6,i)
											strEduStatus = arrEdu(7,i)
											strEduLocal = arrEdu(8,i)
										End If
									%>
									<!-- <tr>
										<td class="alignC"><%=strEduDtB%></td>
										<td class="alignC"><%=strEduDtE%></td>
										<td><%=strEduNm%></td>
										<td class="alignC"><%=strEduMajor%></td>
										<td class="alignC"><%=strEduStatus%></td>
										<td class="alignC"><%=strEduLocal%></td>
									</tr> -->
									<% 
									Next 
									%>
								</tbody>
							</table>
						</div>
					</div>
					<div class="section">
						<h2>경력사항</h2>
						<div class="tblArea">
							<%
							If IsArray(arrCareer) Then
								arrLen = UBound(arrCareer, 2)
							Else
								arrLen = -1
							End If
							%>
							
							<table class="tblForm">
								<colgroup>
									<col class="w85" />
									<col class="w230" />
									<!-- <col class="w100" />
									<col class="w70" /> -->
									<col class="w150" />
									<col />
									<!-- <col class="w60" /> -->
							</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<!-- <td></td>
										<td></td> -->
										<!-- <td></td> -->
									</tr>
									<tr>
										<th rowspan="<%=(2+arrLen)%>">경력사항</th>
										<th>근무기간</th>
										<th>직장명</th>
										<!-- <th>직위</th>
										<th>부서명</th> -->
										<th>담당업무</th>
										<!-- <th>퇴직사유</th> -->

									</tr>
									<%
									
									If arrLen > -1 Then 
									%>
									<% 		
									'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10)
									iRow = arrLen + 1
									If arrLen <= iRow Then
										iBegin = 0
										iEnd = iRow
									Else
										iBegin = arrLen - iRow
										iEnd = arrLen
									End If

									'response.write  iBegin & "<BR>"
									'response.write  iEnd & "<BR>"
									'response.End 

									For i = iBegin To iEnd - 1
										strCarDt = "" : strCarCNm = "" : strCarPos = "" : strCarDpt = "" : strCarTask = "" : strCarRs = ""
										
										If i <= arrLen Then
											If arrCareer(1,i) <> "" And arrCareer(3,i) <> ""  Then
												If arrCareer(1,i) <> "" And arrCareer(2,i) <> "" Then
													strCarDt = arrCareer(1,i) & "." & arrCareer(2,i)
												End If
												
												strCarDt = strCarDt & " ~ "
												
												If arrCareer(3,i) <> "" And arrCareer(4,i) <> "" Then
													strCarDt = strCarDt & arrCareer(3,i) & "." & arrCareer(4,i)
												End If
											End If
											
											strCarCNm = arrCareer(5,i)
											strCarDpt = arrCareer(6,i)
											strCarPos = arrCareer(7,i)
											strCarTask = arrCareer(8,i)
											strCarRs = arrCareer(9,i)
										End If
									%>
									<tr>
										<td class="alignC">

												<%
												Select Case(arrCareer(11,i))
													Case "0"
														Response.Write "1년미만"
													Case "1"
														Response.Write "1년이상"
													Case "2"
														Response.Write "2년이상"
													Case "3"
														Response.Write "3년이상"
													Case "4"
														Response.Write "4년이상"
												End Select 
												%>
											</td>
											<td class="alignC"><%=arrCareer(5,i)%></td>
											<td class="alignC"><%=arrCareer(8,i)%></td>
									</tr>
									<% Next %>
									<%
									End If 
									%>
								</tbody>
							</table>
							
						</div>
					</div>
					<!-- <div class="section">
						<h2>가족사항</h2>
						<div class="tblArea">
							<%
							If IsArray(arrFam) Then
								arrLen = UBound(arrFam, 2)
							Else
								arrLen = -1
							End If
							%>
							<table class="tblForm">
								<colgroup>
									<col class="w85" />
									<col class="w80" />
									<col class="w80" />
									<col class="w80" />
									<col class="w80" />
									<col />
									<col class="w80" />
								</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr>
										<th rowspan="5">가족사항</th>
										<th>관계</th>
										<th>성명</th>
										<th>연령</th>
										<th>학력</th>
										<th>근무처</th>
										<th>동거유무</th>
									</tr>
									<% 		
									' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
									For i = 0 To 3 
										strFamRel = "" : strFamNm = "" : strFamAge = "" : strFamEdu = "" : strFamJob = "" : strFamWith = ""
										
										If i <= arrLen Then
											strFamRel = arrFam(1,i)
											strFamNm = arrFam(2,i)
											strFamAge = arrFam(3,i)
											strFamEdu = arrFam(4,i)
											strFamJob = arrFam(5,i)
											strFamWith = Fn_SetDefault(arrFam(6,i),True,"유","무")
										End If
									%>
									<tr>
										<td class="alignC"><%=strFamRel%></td>
										<td class="alignC"><%=strFamNm%></td>
										<td class="alignC"><%=strFamAge%></td>
										<td><%=strFamEdu%></td>
										<td><%=strFamJob%></td>
										<td class="alignC"><%=strFamWith%></td>
									</tr>
									<% Next %>
								</tbody>
							</table>
						</div>
					</div> -->
					<div class="infoSign">
						<div class="fl_l">
							<p><br>위 기재사항은 사실과 다름 없음.</p>
						</div>
						<div class="fl_r">
							<p><br><%'=FN_SetDateTimeFormat(RegDate,"YYYY년 MM월 DD일")%> 작성자 : <%=UsrName%></p>
						</div>
					</div>
				</div>
			</div>
			
			<div class="section printSec pb50">
				<h2>자기소개서</h2>
				<div class="txtArea"><%=FN_NLToBr(Intro)%></div>
			</div>
			
			<% If RecrSeq > 0 AND IsArray(arrQues) Then ' 사전질문 노출 결정 Start  %>
			<div class="section printSec pb50" id="divQuestion">
				<h2>사전질문</h2>
				<div class="tblArea">
					<%
					arrLen = UBound(arrQues, 2)
					%>
					<table class="tblForm">
						<colgroup>
							<col class="w80" />
							<col />
						</colgroup>
						<tbody>
							<tr class="design">
								<td></td>
								<td></td>
							</tr>
							<% 		
							' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4), Answer(5)
							For i = 0 To arrLen
							%>
							<tr>
								<th>질문 <%=i+1%>.</th>
								<td><%=arrQues(3,i)%></td>
							</tr>
							<tr class="last2">
								<td colspan="2"><%=FN_NLToBr(arrQues(5,i))%></td>
							</tr>
							<% 
							Next 
							%>
						</tbody>
					</table>
				</div>
			</div>
			<% End If ' 사전질문 노출 결정 End %>

		</div>

<%
			End If
		Next
	End Sub	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>이력서 인쇄 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<style>
		@media print {
			.printSec {page-break-before:always;}
			.btns {display:none;}
			#contArea h2 {font-size:22px; height:25px;line-height:25px; padding-top:25px;}
		}
		#divQuestion .tblForm td {line-height:18px;}
	</style>
	</head>
	<script type="text/javascript">
	<!--
		$(function() {
			$(".btnPrint").click(function(){ window.print(); });
		});
	//-->
	</script>
</head>
<body class="pop">
	<div id="wrap" class="infoPrint">
		<% Call Main() %>
	</div>
</body>
</html>