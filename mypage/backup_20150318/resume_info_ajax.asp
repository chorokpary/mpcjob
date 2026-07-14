<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : resume_info_ajax.asp / 회원상세정보 (ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-24 / 강미현 / 4M / 학력 등록 없을 경우 최종학력 노출 X
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

		Seq = Session("USeq")
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
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort")
			Case "FAM" :
				TargetData = "uspUserFam_Info"
				kFields 	= Array("UsrSeq","Relation","FmName","Age","Education","Job","IsLiveWith","Sort")
			Case "MEMO" :
				TargetData = "uspUserMemo_Info"
				kFields 	= Array("UsrMSeq","AdmName","Memo")
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
						<% If arrLen > -1 Then %>
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr>
									<th>최종학력</th>
									<td class="alignL"> <strong><% If arrLen > -1 Then %> <%=strFinEdu%> <% End If %></strong></th>
								</tr>
							</tbody>
						</table>
						<% End If %>
						<table class="tblForm alignC">
							<colgroup>
								<col class="w30" />
								<col class="w140" />
								<col class="w140" />
								<col />
								<col class="w120" />
								<col class="w90" />
								<col class="w100" />
							</colgroup>
							<tbody>
								
								<tr>
									<th></th>
									<th>입학년/월</th>
									<th>졸업년/월</th>
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
									<td><input type="checkbox" id="EduSort_<%=arrData(9,i)%>" name="EduSort" value="<%=arrData(9,i)%>" title="선택" /></td>
									<td><% If arrData(1,i) <> "" And arrData(2,i) <> "" Then %> <%=arrData(1,i)%>년 <%=arrData(2,i)%>월 <% End If %></td>
									<td><% If arrData(3,i) <> "" And arrData(4,i) <> "" Then %> <%=arrData(3,i)%>년 <%=arrData(4,i)%>월 <% End If %></td>
									<td><%=arrData(5,i)%></td>
									<td><%=arrData(6,i)%></td>
									<td><%=arrData(7,i)%></td>
									<td><%=arrData(8,i)%></td>
								</tr>
								<% Next %>
								<tr>
									<td></td>
									<td>
										<select id="EduBYear" name="EduBYear" title="입학년도" class="w55">
											<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
											<option value="<%=yy%>"><%=yy%></option>
											<% Next %>
										</select> 년
										<select id="EduBMonth" name="EduBMonth" title="입학월" class="w40">
											<% For mm = 1 TO 12 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select> 월
									</td>
									<td>
										<select id="EduEYear" name="EduEYear" title="졸업년도" class="w55">
											<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
											<option value="<%=yy%>"><%=yy%></option>
											<% Next %>
										</select> 년
										<select id="EduEMonth" name="EduEMonth" title="졸업월" class="w40">
											<% For mm = 1 TO 12 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select> 월
									</td>
									<td><input type="text" id="SchoolName" name="SchoolName" title="학교명" value="" class="w120" maxlength="20" /></td>
									<td><input type="text" id="Major" name="Major" title="전공" value="" class="w100" maxlength="10" /></td>
									<td>
										<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Status", "Status", "", "", "", "", " class=""w65"" ", False)%>
									</td>
									<td>
										<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Local", "Local", "", "", "", "", " class=""w75"" ", False)%>
									</td>
								</tr>
							</tbody>
						</table>
<% ElseIf Flag = "CAREER" Then %>
						<table class="tblForm alignC">
							<colgroup>
								<col class="w30" />
								<col />
								<col class="w125" />
								<col class="w50" />
								<col class="w80" />
								<col class="w145" />
								<col class="w70" />
							</colgroup>
							<tbody>
								<tr>
									<th></th>
									<th>근무기간</th>
									<th>직장명</th>
									<th>직위</th>
									<th>부서명</th>
									<th>담당업무</th>
									<th>퇴직사유</th>
								</tr>
								<% 		
								'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10)
								For i = 0 To arrLen 
								%>
								<tr>
									<td class="alignC"><input type="checkbox" id="CareerSort_<%=arrData(10,i)%>" name="CareerSort" value="<%=arrData(10,i)%>" title="선택" /></td>
									<td class="alignC">
										<% If arrData(1,i) <> "" And arrData(3,i) <> ""  Then %>
											<% If arrData(1,i) <> "" And arrData(2,i) <> "" Then %><%=arrData(1,i)%>.<%=arrData(2,i)%><% End If %>
											~
											<% If arrData(3,i) <> "" And arrData(4,i) <> "" Then %><%=arrData(3,i)%>.<%=arrData(4,i)%><% End If %>
										<% End If %>
									</td>
									<td><%=arrData(5,i)%></td>
									<td class="alignC"><%=arrData(7,i)%></td>
									<td class="alignC"><%=arrData(6,i)%></td>
									<td><%=arrData(8,i)%></td>
									<td><%=arrData(9,i)%></td>
								</tr>
								<% Next %>
								<tr>
									<td class="alignC"></td>
									<td class="alignC">
										<select id="CrBYear" name="CrBYear" title="입사년도" class="w50">
											<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
											<option value="<%=yy%>"><%=yy%></option>
											<% Next %>
										</select>년
										<select id="CrBMonth" name="CrBMonth" title="입사월" class="w40">
											<% For mm = 1 TO 12 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select>월
										~
										<select id="CrEYear" name="CrEYear" title="퇴사년도" class="w50">
											<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
											<option value="<%=yy%>"><%=yy%></option>
											<% Next %>
										</select>년
										<select id="CrEMonth" name="CrEMonth" title="퇴사월" class="w40">
											<% For mm = 1 TO 12 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select>월
									</td>
									<td><input type="text" id="CompanyName" name="CompanyName" title="직장명" value="" class="w100" maxlength="20" /></td>
									<td class="alignC"><input type="text" id="Position" name="Position" title="직위" value="" class="w30" maxlength="20" /></td>
									<td class="alignC"><input type="text" id="Department" name="Department" title="부서" value="" class="w60" maxlength="15" /></td>
									<td><input type="text" id="Task" name="Task" name="Task" title="업무" value="" class="w125" maxlength="15" /></td>
									<td><input type="text" id="Reason" name="Reason" name="Reason" title="퇴사사유" value="" class="w50" maxlength="25" /></td>
								</tr>
							</tbody>
						</table>
<% ElseIf Flag = "FAM" Then %>

						<table class="tblForm alignC">
							<colgroup>
								<col class="w30" />
								<col class="w115" />
								<col class="w155" />
								<col class="w75" />
								<col class="w115" />
								<col />
								<col class="w100" />
							</colgroup>
							<tbody>
								<tr>
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
<% End If %>
