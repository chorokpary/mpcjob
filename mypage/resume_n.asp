<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : resume.asp / 회원정보 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2014-01-07 / 강미현 / 4M / 희망분야 등록항목 삭제
'   #003 : 2014-12-16 / 이길홍 / 4M / 주민등록 등록항목 삭제
'	#004 : 2015-02-24 / 이길홍 / 4M / 폼 변경 
'____________________________________________________________________________________
%>
<%
	Dim Seq, Page
	Dim UsrName, Birthday, Gender, Mobile, Email, HopePart, HopePart2, UsrType
	Dim Zipcode1, Zipcode2, AddrSido, AddrGugun, Addr, AddrDetail, Jumin1, Jumin2
	Dim ImgPath, Attach1_Path, Attach2_Path, Attach3_Path, Attach1_FName, Attach2_FName, Attach3_FName
	Dim RegDate, ModDate, Intro, EduLast, EduType
	Dim kFields, arrLen, i, yy, mm
	Dim arrEdu, arrCareer, arrFam, arrMemo, arrApply

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)
	

	Call SB_LoginChkPage()
	Call Main()
	
	Sub Main()

		Seq = Session("USeq")
		Page = FN_Req("Page","1")
		
		If Seq <> "0" Then
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs

		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				UsrName 	= Rs("UsrName")
				Birthday 	= Rs("Birthday")
				Gender 	= Rs("Gender")
				Mobile 	= oSeed.Decrypt(Rs("Mobile"))
				Email 	= Rs("Email")
				HopePart 	= Rs("HopePart")
				HopePart2 = Rs("HopePart2")
				UsrType 	= Rs("UsrType")
				Zipcode1 = Rs("Zipcode1")
				Zipcode2 = Rs("Zipcode2")
				AddrSido 	= Rs("AddrSido")
				AddrGugun = Rs("AddrGugun")
				Addr 	= Rs("Addr")
				AddrDetail = Rs("AddrDetail")
				Jumin1 	= Rs("Jumin1")
				Jumin2 	= Rs("Jumin2")

				ImgPath 	= Rs("ImgPath")
				Attach1_Path = Rs("Attach1_Path")
				Attach1_FName = Rs("Attach1_FName")
				Attach2_Path 	= Rs("Attach2_Path")
				Attach2_FName = Rs("Attach2_FName")
				Attach3_Path 	= Rs("Attach3_Path")
				Attach3_FName = Rs("Attach3_FName")
				
				RegDate = Rs("RegDate")
				ModDate = Rs("ModDate")
				
				Intro = FN_NLToBr(Rs("Intro"))

				EduLast = Rs("EduLast")
				EduType = Rs("EduType")

			End If

			Set Rs = Rs.NextRecordSet()
				
			'arrDataNum 	= UBound(arrData, 2)
			If Not Rs.Eof Then
				' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
				kFields 	= Array("UsrSeq","EduBYear","EduBMonth","EduEYear","EduEMonth","SchoolName","Major","Status","Local","Sort")
				arrEdu 	= Rs.GetRows(,,kFields)
			End If
			
			Set Rs = Rs.NextRecordSet()
				
			If Not Rs.Eof Then
				' UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10), CompanyYear(11)
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort" ,"CompanyYear")
				arrCareer 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
				kFields 	= Array("UsrSeq","Relation","FmName","Age","Education","Job","IsLiveWith","Sort")
				arrFam 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrMSeq(0), AdmName(1), Memo(2)
				kFields 	= Array("UsrMSeq","AdmName","Memo")
				arrMemo 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' PrjName(0), Title(1), AppBDate(2), AppEDate(3), RegDate(4), AdmName(5), CdVal(6)
				kFields 	= Array("PrjName","Title","AppBDate","AppEDate","RegDate","AdmName","CdVal")
				arrApply 	= Rs.GetRows(,,kFields)
			End If			

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>MY PAGE - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function goChangeMobile() {
			openLayer("change_mobile.asp","", "w528");
		}
		
		function goChangePasswd() {
			openLayer("change_passwd.asp","", "w528");
		}
	</script>
</head>
<body>
	<div id="wrap" class="resumeMod">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_resumeMng.jpg" alt="이력서관리" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<h3><img src="/images/pages/stit_resumeMod_01.jpg" alt="개인정보" /></h3>
				<p class="note"> <strong><%=UsrName%></strong> 회원님의 기본정보입니다. </p>
				<div class="section">
					<div class="imgs">
						<div class="inner">
							<% If Not FN_isBlank(ImgPath) Then %>
							<img src="<%=ImgPath%>" width="114" height="144">
							<% Else %>
							<img src="/images/pages/img_nopics.jpg" alt="" />
							<% End If %>
						</div>
					</div>	
					<div class="info">
						<table>
							<caption>이력서 개인정보</caption>
							<colgroup>
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr>
									<th><label for="">이름</label></th>
									<td><%=UsrName%></td>
									<th><label for="">나이</label></th>
									<td><% If IsNumeric(Left(Birthday,4)) Then %><%=Year(Date) - Left(Birthday,4) + 1%> 세<% Else %> - <% End If %></td>
								</tr>
								<tr>
									<th><label for="">성별</label></th>
									<td colspan="3">
										<%
											Select Case Gender 
												Case "1" : Response.Write "남" 
												Case "2" : Response.Write "여" 
												Case Else : Response.Write "-" 
											End Select 
										%>									
									</td>
									<!--
									<th><label for="">희망분야</label></th>
									<td><%=HopePart%> <% If HopePart2 <> "" AND Not ISNULL(HopePart2) Then Response.Write " / " & HopePart2 %></td>
									//-->
								</tr>
								<!-- <tr>
									<th><label for="">주민등록번호</label></th>
									<td colspan="3"><%=Jumin1%><%=Fn_SetDefault(Jumin2,"","","-*******")%></td>
								</tr> -->
								<tr>
									<th><label for="">휴대폰</label></th>
									<td colspan="3"><%=Mobile%>
										<a href="javascript:goChangeMobile();"><img src="/images/pages/btn_phoneMod.jpg" alt="휴대폰번호 변경" /></a>
									</td>
								</tr>
								<tr>
									<th><label for="">주소</label></th>
									<td colspan="3"><%=AddrSido%>&nbsp;<%=AddrGugun%></td>
								</tr>
								<tr>
									<th><label for="">E-mail</label></th>
									<td colspan="3"><%=Email%></td>
								</tr>
							</tbody>
						</table>
					</div>
					<div class="alignC mt20"> <a href="javascript:goChangePasswd();"><img src="/images/pages/btn_passMod.jpg" alt="비밀번호 변경" /></a> </div>
				</div>
				
				<h3><img src="/images/pages/stit_resumeMod_02.jpg" alt="학력사항" /></h3>
				<div class="section">
					<%
					Dim strFinEdu : strFinEdu = ""
					If IsArray(arrEdu) Then
						arrLen = UBound(arrEdu, 2)
						strFinEdu = arrEdu(5,arrLen)
					Else
						arrLen = -1
					End If
					%>
					<% If arrLen > -1 Then %>
					<!-- <table>
						<caption>이력서 최종학력</caption>
						<colgroup>
							<col class="w100" />
							<col />
						</colgroup>
						<tbody>
							<tr>
								<th><label for="">최종학력</label></th>
								<td><strong><%=strFinEdu%></strong> </td>
							</tr>
						</tbody>
					</table> -->
					<% End If %>
					<table class="cell_alignC">
						<caption>이력서 개인정보</caption>
						<colgroup>
							<col class="w150" />
							<col />
						</colgroup>
						<!-- <thead>
							<tr>
								<th><label for="">입학년/월</label></th>
								<th><label for="">졸업년/월</label></th>
								<th><label for="">학교명</label></th>
								<th><label for="">전공</label></th>
								<th><label for="">졸업여부</label></th>
								<th><label for="">소재지</label></th>
							</tr>
						</thead> -->
						<tbody>
							<% 		
							' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
							For i = 0 To arrLen 
							%>
							<!-- <tr>
								<td><% If arrEdu(1,i) <> "" And arrEdu(2,i) <> "" Then %> <%=arrEdu(1,i)%>/<%=arrEdu(2,i)%><% End If %></td>
								<td><% If arrEdu(3,i) <> "" And arrEdu(4,i) <> "" Then %> <%=arrEdu(3,i)%>/<%=arrEdu(4,i)%><% End If %></td>
								<td><%=arrEdu(5,i)%></td>
								<td><%=arrEdu(6,i)%></td>
								<td><%=arrEdu(7,i)%></td>
								<td><%=arrEdu(8,i)%></td>
							</tr> -->
							<% Next %>

							<tr>
								<td>
									<!-- <select id="EduLast" name="EduLast" title="최종학력" class="w120">
										<option value="0" <%=Fn_SetDefault(EduLast,"0"," selected","")%>>중학교</option>
										<option value="1" <%=Fn_SetDefault(EduLast,"1"," selected","")%>>고등학교</option>
										<option value="2" <%=Fn_SetDefault(EduLast,"2"," selected","")%>>대학(2,3년)</option>
										<option value="3" <%=Fn_SetDefault(EduLast,"3"," selected","")%>>대학교(4년)</option>
										<option value="4" <%=Fn_SetDefault(EduLast,"4"," selected","")%>>대학원</option>
									</select> -->
									<%
											Select Case EduLast 
												Case "0" : Response.Write "중학교" 
												Case "1" : Response.Write "고등학교" 
												Case "2" : Response.Write "대학(2,3년)" 
												Case "3" : Response.Write "대학교(4년)" 
												Case "4" : Response.Write "대학원" 
											End Select 
										%>	
								</td>
								<td class="alignL">
									<!-- &nbsp;<input type="radio" name="EduType" id="EduType1" value="1" <%=Fn_SetDefault(EduType,"1"," checked","")%> >졸업
									<input type="radio" name="EduType" id="EduType2" value="2" <%=Fn_SetDefault(EduType,"2"," checked","")%>>재학중
									<input type="radio" name="EduType" id="EduType3" value="3" <%=Fn_SetDefault(EduType,"3"," checked","")%>>휴학
									<input type="radio" name="EduType" id="EduType4" value="4" <%=Fn_SetDefault(EduType,"4"," checked","")%>>중퇴 -->
									<%
											Select Case EduType 
												Case "1" : Response.Write "졸업" 
												Case "2" : Response.Write "재학중" 
												Case "3" : Response.Write "휴학" 
												Case "4" : Response.Write "중퇴" 
											End Select 
									%>	
								</td>
							</tr>

						</tbody>
					</table>
				</div>
				
				
				<h3><img src="/images/pages/stit_resumeMod_03.jpg" alt="경력사항" /></h3>
				<div class="section">
					<%
					If IsArray(arrCareer) Then
						arrLen = UBound(arrCareer, 2)
					Else
						arrLen = -1
					End If
					%>
					<table class="cell_alignC">
						<caption>이력서 개인정보</caption>
						<colgroup>
							<col class="w100" />
							<col class="w80" />
							<col class="w80" />
							<col class="w200" />
							<col class="w100" />
							<col />
						</colgroup>
						<!-- <thead>
							<tr>
								<th><label for="">근무기간</label></th>
								<th><label for="">직장명</label></th>
								<th><label for="">직위</label></th>
								<th><label for="">부서명</label></th>
								<th><label for="">담당업무</label></th>
								<th><label for="">퇴직사유</label></th>
							</tr>
						</thead> -->

						<tbody>
							<% 		
							'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10)
							For i = 0 To arrLen 
							%>
							<tr>
									<td class="alignC">근무개월수</td>
									<td>
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
									<td class="alignC">회사명</td>
									<td class="alignC"><%=arrCareer(5,i)%></td>
									<td>담당업무</td>
									<td><%=arrCareer(8,i)%></td>
							</tr>
							<% Next %>
						</tbody>
					</table>
				</div>

				<!-- <h3><img src="/images/pages/stit_resumeMod_04.jpg" alt="가족사항" /></h3>
				<div class="section">
					<%
					If IsArray(arrFam) Then
						arrLen = UBound(arrFam, 2)
					Else
						arrLen = -1
					End If
					%>
					<table class="cell_alignC">
						<caption>이력서 개인정보</caption>
						<colgroup>
							<col class="w115" />
							<col class="w165" />
							<col class="w100" />
							<col class="w100" />
							<col />
							<col class="w100" />
						</colgroup>
						<thead>
							<tr>
								<th><label for="">관계</label></th>
								<th><label for="">성명</label></th>
								<th><label for="">연령</label></th>
								<th><label for="">학력</label></th>
								<th><label for="">근무처</label></th>
								<th><label for="">동거유무</label></th>
							</tr>
						</thead>
						<tbody>
							<% 		
							' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
							For i = 0 To arrLen 
							%>
							<tr>
								<td><%=arrFam(1,i)%></td>
								<td><%=arrFam(2,i)%></td>
								<td><%=arrFam(3,i)%></td>
								<td><%=arrFam(4,i)%></td>
								<td><%=arrFam(5,i)%></td>
								<td><%=Fn_SetDefault(arrFam(6,i),True,"유","무")%></td>
							</tr>
							<% Next %>
						</tbody>
					</table>
				</div> -->
				<h3><img src="/images/pages/stit_resumeMod_05.jpg" alt="자기소개서" /></h3>
				<div class="textAreas"><%=Intro%></div>
				<div class="alignC mt30"> <a href="/recruit/recruit_list.asp"><img src="/images/pages/btn_goApply.jpg" alt="입사지원하러가기" /></a>
					<a href="resume_modify.asp"><img src="/images/pages/btn_modify.jpg" alt="수정하기" /></a></div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>