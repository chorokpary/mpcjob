<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_info_ajax.asp / 회원상세정보 ajax
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-19 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag
	Dim kFields, arrLen, i, yy, mm
	Dim arrData
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Flag = FN_Req("Flag","")
		
		If Seq <> "0" Then
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		Dim TargetData
		
		Select Case Flag
			Case "EDU" :
				TargetData = "uspUserEdu_Info"
				kFields 	= Array("UsrSeq","EduBYear","EduBMonth","EduEYear","EduEMonth","SchoolName","Major","Status","Local","Sort")
			Case "CAREER" :
				TargetData = "uspUserCareer_Info"
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort", "CompanyYear")
			Case "FAM" :
				TargetData = "uspUserFam_Info"
				kFields 	= Array("UsrSeq","Relation","FmName","Age","Education","Job","IsLiveWith","Sort")
			Case "MEMO" :
				TargetData = "uspUserMemo_Info"
				kFields 	= Array("UsrMSeq","AdmName","Memo","RegDate")
			Case Else :
				Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","BACK")
		End Select
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = TargetData
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "M_ALL"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				arrData = Rs.GetRows(,,kFields)
				arrLen 	= UBound(arrData, 2)
			Else
				arrLen = -1
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<% If Flag = "EDU" Then %>
<%
		Dim strFinEdu
		
		If arrLen >= 0 Then
			strFinEdu = arrData(5,arrLen)
		End If 
%>

								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col class="w30" />
										<col class="w110" />
										<col class="w110" />
										<col />
										<col class="w100" />
										<col class="w80" />
										<col class="w100" />
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
											<td></td>
										</tr>
										<tr>
											<th rowspan="<%=(4+arrLen)%>">학력사항</th>
											<th colspan="7" class="alignL"><% If arrLen > -1 Then %>최종학력 : <%=strFinEdu%> <% End If %></th>
										</tr>
										<tr>
											<th></th>
											<th>입학년월</th>
											<th>졸업년월</th>
											<th>학교명</th>
											<th>전공</th>
											<th>졸업여부</th>
											<th>소재지</th>
										</tr>
										<% 		
										' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
										For i = 0 To arrLen 
										%>
										<tr>
											<td class="alignC"><input type="checkbox" id="EduSort_<%=arrData(9,i)%>" name="EduSort" value="<%=arrData(9,i)%>" title="선택" /></td>
											<td class="alignC"><% If arrData(1,i) <> "" And arrData(2,i) <> "" Then %> <%=arrData(1,i)%> 년 <%=arrData(2,i)%> 월 <% End If %></td>
											<td class="alignC"><% If arrData(3,i) <> "" And arrData(4,i) <> "" Then %> <%=arrData(3,i)%> 년 <%=arrData(4,i)%> 월 <% End If %></td>
											<td><%=arrData(5,i)%></td>
											<td class="alignC"><%=arrData(6,i)%></td>
											<td class="alignC"><%=arrData(7,i)%></td>
											<td class="alignC"><%=arrData(8,i)%></td>
										</tr>
										<% Next %>
										<tr>
											<td class="alignC"></td>
											<td class="alignC">
												<select id="EduBYear" name="EduBYear" title="입학년도" class="w55">
													<% For yy = 1970 TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="EduBMonth" name="EduBMonth" title="입학월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
											</td>
											<td class="alignC">
												<select id="EduEYear" name="EduEYear" title="졸업년도" class="w55">
													<% For yy = 1970 TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="EduEMonth" name="EduEMonth" title="졸업월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
											</td>
											<td><input type="text" id="SchoolName" name="SchoolName" title="학교명" value="" class="w145" maxlength="20" /></td>
											<td class="alignC"><input type="text" id="Major" name="Major" title="전공" value="" class="w80" maxlength="10" /></td>
											<td class="alignC">
												<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Status", "Status", "", "", "", "", " class=""w55"" ", False)%>
											</td>
											<td class="alignC">
												<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Local", "Local", "", "", "", "", " class=""w55"" ", False)%>
											</td>
										</tr>
									</tbody>
								</table>
<% ElseIf Flag = "CAREER" Then %>
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col class="w30" />
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
											<td></td>
											<!-- <td></td> -->
										</tr>
										<tr>
											<th rowspan="<%=(3+arrLen)%>">경력사항</th>
											<th></th>
											<th>근무기간</th>
											<th>직장명</th>
											<!-- <th>직위</th>
											<th>부서명</th> -->
											<th>담당업무</th>
											<!-- <th>퇴직사유</th> -->
										</tr>
										<% 		
										'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10)
										For i = 0 To arrLen 
										%>
										<tr>
											<td class="alignC"><input type="checkbox" id="CareerSort_<%=arrData(10,i)%>" name="CareerSort" value="<%=arrData(10,i)%>" title="선택" /></td>
											<td class="alignC">
												<!-- <% If arrData(1,i) <> "" And arrData(3,i) <> ""  Then %>
													<% If arrData(1,i) <> "" And arrData(2,i) <> "" Then %><%=arrData(1,i)%>.<%=arrData(2,i)%><% End If %>
													~
													<% If arrData(3,i) <> "" And arrData(4,i) <> "" Then %><%=arrData(3,i)%>.<%=arrData(4,i)%><% End If %>
												<% End If %> -->
												<%
												Select Case(arrData(11,i))
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
											<td><%=arrData(5,i)%></td>
											<!-- <td class="alignC"><%=arrData(7,i)%></td>
											<td class="alignC"><%=arrData(6,i)%></td> -->
											<td><%=arrData(8,i)%></td>
											<!-- <td><%=arrData(9,i)%></td> -->
										</tr>
										<% Next %>
										<tr>
											<td class="alignC"></td>
											<td class="alignC">
												<!-- <select id="CrBYear" name="CrBYear" title="입사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrBMonth" name="CrBMonth" title="입사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
												~
												<select id="CrEYear" name="CrEYear" title="퇴사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrEMonth" name="CrEMonth" title="퇴사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select> -->
												<select id="CompanyYear" name="CompanyYear" title="근무개월수" class="w70">
													<option value="0">1년미만</option>
													<option value="1">1년이상</option>
													<option value="2">2년이상</option>
													<option value="3">3년이상</option>
													<option value="4">4년이상</option>
												</select>
											</td>
											<td><input type="text" id="CompanyName" name="CompanyName" title="직장명" value="" class="w80" maxlength="20" /></td>
											<!-- <td class="alignC"><input type="text" id="Position" name="Position" title="직위" value="" class="w50" maxlength="20" /></td>
											<td class="alignC"><input type="text" id="Department" name="Department" title="부서" value="" class="w60" maxlength="15" /></td> -->
											<td><input type="text" id="Task" name="Task" name="Task" title="업무" value="" class="w105" maxlength="15" /></td>
											<!-- <td><input type="text" id="Reason" name="Reason" name="Reason" title="퇴사사유" value="" class="w45" maxlength="25" /></td> -->
										</tr>
									</tbody>
								</table>
<% ElseIf Flag = "FAM" Then %>
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col class="w30" />
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
											<td></td>
										</tr>
										<tr>
											<th rowspan="<%=(3+arrLen)%>">가족사항</th>
											<th></th>
											<th>관계</th>
											<th>성명</th>
											<th>연령</th>
											<th>학력</th>
											<th>근무처</th>
											<th>동거유무</th>
										</tr>
										<% 		
										' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
										For i = 0 To arrLen 
										%>
										<tr>
											<td class="alignC"><input type="checkbox" id="FamSort_<%=arrData(7,i)%>" name="FamSort" value="<%=arrData(7,i)%>" title="선택" /></td>
											<td class="alignC"><%=arrData(1,i)%></td>
											<td class="alignC"><%=arrData(2,i)%></td>
											<td class="alignC"><%=arrData(3,i)%></td>
											<td><%=arrData(4,i)%></td>
											<td><%=arrData(5,i)%></td>
											<td class="alignC"><%=Fn_SetDefault(arrData(6,i),True,"유","무")%></td>
										</tr>
										<% Next %>
										<tr>
											<td class="alignC"></td>
											<td class="alignC"><input type="text" id="Relation" name="Relation" title="관계" value="" class="w60" maxlength="5" /></td>
											<td class="alignC"><input type="text" id="FmName" name="FmName" title="성명" value="" class="w60" maxlength="10" /></td>
											<td class="alignC"><input type="text" id="Age" name="Age" title="연령" value="" class="w60" maxlength="3" /></td>
											<td><input type="text" id="Education" name="Education" title="학력" value="" class="w60" maxlength="15" /></td>
											<td><input type="text" id="Job" name="Job" title="근무처" value="" class="w245" maxlength="15" /></td>
											<td class="alignC">
											<select id="IsLiveWith" name="IsLiveWith" class="w60" title="동거유무">
												<option value="">:유/무:</option>
												<option value="1">유</option>
												<option value="0">무</option>
											</select>
											</td>
										</tr>
									</tbody>
								</table>
<% ElseIf Flag = "MEMO" Then %>
							<table>
								<colgroup>
									<col class="w80" />
									<col />
									<col class="w150" />
									<col class="w60" />
								</colgroup>
								<thead>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr>
										<th>작성자</th>
										<th>내용</th>
										<th>작성일시</th>
										<th>삭제</th>
									</tr>
								</thead>
								<tbody>
									<% 	
									If IsArray(arrData) Then 	
										' UsrMSeq(0), AdmName(1), Memo(2), RegDate(3)
										For i = 0 To UBound(arrData, 2) 
									%>
									<tr>
										<td><%=arrData(1,i)%></td>
										<td class="alignL"><%=arrData(2,i)%></td>
										<td class="alignL"><%=arrData(3,i)%></td>
										<td><a href="javascript:delDetailInfo('Memo',<%=arrData(0,i)%>)"><img src="/admin/images/btn_del.jpg" alt="삭제" /></a></td>
									</tr>
									<% 
										Next 
									Else
									%>
									<tr>
										<td colspan="4">등록된 메모가 없습니다.</td>
									</tr>
									<%
									End If
									%>
								</tbody>
							</table>
<% End If %>
