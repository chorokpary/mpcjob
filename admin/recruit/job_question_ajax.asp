<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_question_ajax.asp / 회원상세정보 ajax
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-19 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Token, Flag
	Dim kFields, arrLen, i, yy, mm
	Dim arrData
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Token = FN_Req("Token","")
		Flag = FN_Req("Flag","")
		
		If Seq <> "0" Then
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		Dim TargetData
		
		Select Case Flag
			Case "QUE" :
				TargetData = "uspRecruitQaQ_Info"
				kFields 	= Array("QSeq","QCode","Sort","Question","IsDisp")
			Case Else :
				Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","BACK")
		End Select
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = TargetData
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "LST_ADM"
			.Parameters("@Seq").Value 	= Seq
			.Parameters("@Token").Value = Token
			
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
<% If Flag = "QUE" Then %>
									<% 		
									' ' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4)
									Dim Num
									For i = 0 To arrLen 
										Num = i + 1
									%>
									<tr class="list">
										<td class="alignC"><input type="text" name="QueSort_<%=Num%>" id="QueSort_<%=Num%>" title="노출순서" value="<%=FN_InputVal(arrData(2,i))%>" class="w50 alignC num" maxlength="3" /></td>
										<td class="alignC"><input type="text" name="Question_<%=Num%>" id="Question_<%=Num%>" value="<%=FN_InputVal(arrData(3,i))%>" class="w530" maxlength="200"></td>
										<td class="alignC">
											<input type="radio" name="QueIsDisp_<%=Num%>" id="QueIsDisp_<%=Num%>_Y" value="1" <%=Fn_SetDefault(arrData(4,i), "True", " checked", "")%>><label for="QueIsDisp_<%=Num%>_Y">노출</label>
											<input type="radio" name="QueIsDisp_<%=Num%>" id="QueIsDisp_<%=Num%>_N" value="0" <%=Fn_SetDefault(arrData(4,i), "False", " checked", "")%>> <label for="QueIsDisp_<%=Num%>_N">숨김</label>
										</td>
										<td>
											<span class="btns btn_gr"><a href="javascript:delDetailInfo('Que',<%=arrData(0,i)%>)">삭제</a></span>
											<input type="hidden" id="QueSeq_<%=Num%>" name="QueSeq_<%=Num%>" value="<%=arrData(0,i)%>" />
										</td>
										<% If i = 0 Then %>
										<td rowspan="<%=arrLen+1%>">
											<span class="btns btn_b"><a href="javascript:addDetailInfo('Que',<%=arrLen+1%>)">수정</a></span>
										</td>
										<% End If %>
									</tr>
									<% Next %>
<% End If %>
