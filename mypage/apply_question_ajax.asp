<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : apply_question_ajax.asp / 입사지원 - 사전질문 작성폼 & 데이터 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2014-09-25 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim RecrSeq, PassQues, DtSize

	Dim arrQues, arrQuesLen
	
	Dim Result
	
	Call Main()
	
	Sub Main()
		Dim fnMsg
	
		RecrSeq 	= FN_Req("RecrSeq","0")
		PassQues = FN_Req("PassQues","N")	' 사전질문 답변 여부 체크
		DtSize = FN_Req("DtSize","0")
		
		If PassQues <> "Y" Then	PassQues = "N"
			
			
		If RecrSeq <> "0" Then
			If PassQues = "N" Then
				' 사전질문 여부 확인
				Call getApplyList()
			End If

			If PassQues = "N" Then
				' 사전질문 작성폼
				Call setApplyQuestion()
				Response.End
			Else
				' 사전질문 답변 등록
				Dim fnQCode, fnAnswer, fnI
				
				If IsNumeric(DtSize) Then
					For fnI = 1 To DtSize
						fnQCode = FN_Req("QCode_"& fnI ,"0")
						fnAnswer = FN_Req("Answer_"& fnI ,"")
						
						If Not IsNumeric(fnQCode) Then	fnQCode = 0
						If fnQCode > 0 Then
							Call setApplyAnswer(fnQCode, fnAnswer)
						End If
					Next
				End If
			End If
		Else
			Result = 99
		End If
		
		If Result = 0 Then
			fnMsg = "사전질문 등록이 완료되었습니다."
		Else
			fnMsg = "사전질문 등록중 오류가 발생하였습니다."
		End If
		Call setResult(fnMsg)
		
	End Sub


	Sub setResult(argMsg)
		Dim strMsg, strScript
		Dim strUrl
		
		%>
			<div class="inner"><div class="inner2">
				<div class="popCont question">
					<h2><img src="/images/pages/tit_applyQuestion.jpg" alt="사전질문" /></h2>
					<div class="popCont message ">
					
						<div class="findPop msg"><p><%=argMsg%></p></div>
						<p class="popBtn mt20">
							<a href="javascript:;" id="btnClose"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
						</p>
					</div>
				</div>
			</div></div>
		<%
		Response.End

	End Sub
	
	' #### 사전질문
	Sub setApplyAnswer(argQCode, argAnswer)
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruitQaA_Proc"
			.Parameters.Refresh
			.Parameters("@RecrSeq").Value 		= RecrSeq
			.Parameters("@QCode").Value 		= argQCode
			.Parameters("@Answer").Value 		= argAnswer
			.Parameters("@UsrSeq").Value 		= Session("USeq")
			
			.Execute
			
			Result = .Parameters("@Result")
			
		End With
	End Sub
	
	Sub getApplyList()
		Dim Rs, kFields
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = "uspRecruitQaQ_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "LST_FRONT"
			.Parameters("@Seq").Value 	= RecrSeq
			.Parameters("@UsrSeq").Value = Session("USeq")
			
			' ## 회원정보
			Set Rs = .Execute
	
			If Not Rs.Eof Then
				kFields 	= Array("QSeq","QCode","Sort","Question","IsDisp","Answer")
				arrQues = Rs.GetRows(,,kFields)
				arrQuesLen 	= UBound(arrQues, 2)
			Else
				'arrQuesLen = -1
				' 사전질문이 없을경우 창 닫음		
				Call setResult("사전질문을 등록하지 않은 채용공고 입니다.")
				Response.End
			End If
	
			Rs.Close
			Set Rs = Nothing
		End With	
	End Sub
	
	' 사전질문 폼
	Sub setApplyQuestion()
		Dim fnI : fnI = 0
		Dim Num
		%>
			<script type="text/javascript">
				function chkQuesForm(){
					var len = $("#DtSize").val();
					var ckForm = true;
					
					if (len==0){
						alert("잘못된 경로로 접근하셨습니다.");
						closeLayer();
					}else{
						/*
						// 유효성체크
						$("#frmQuestion td.a textarea").each(function(){
							if (!$(this).val()){
								alert("질문의 답변을 입력해주세요.");
								$(this).focus();
								ckForm = false;
								return false;
							}
						});
						*/
					}
					
					if (ckForm) {
						var url = "/mypage/apply_question_ajax.asp"
						param = $("#frmQuestion").serialize();
						
						$.ajax({
							type: "POST"
							, url: url
							, data: param
							, success: function(content){ 
								$("#divPopArea .popBox").html(content);
							}
							, error: function(xhr, status, error) { alert("[입사지원] " + error);}
						});
					}
				}
			</script>
			<div class="inner"><div class="inner2">
				<div class="popCont question">
					<h2><img src="/images/pages/tit_applyQuestion.jpg" alt="사전질문" /></h2>
					<!-- p class="note">
						<img src="/images/pages/txt_xxx.jpg" alt="사전질문 관련 안내" /><br />
					</p //-->
					
					<form id="frmQuestion" name="frmQuestion" method="post" >
						<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
						<input type="hidden" id="PassQues" name="PassQues" value="Y">
						<input type="hidden" id="DtSize" name="DtSize" value="<%=arrQuesLen+1%>">
					
						<div class="board_write mt20">
							<table>
								<caption>사전질문양식</caption>
								<colgroup>
									<col class="w80" />
									<col />
								</colgroup>
								<tbody>
									<% 		
									' ' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4), Answer(5)
									For fnI = 0 To arrQuesLen
										Num = fnI + 1
									%>
									<tr>
										<th><span>질문 <%=Num%></span></th>
										<td class="q"><%=FN_InputVal(arrQues(3,fnI))%></td>
									</tr>
									<tr>
										<td colspan="2" class="a">
											<textarea name="Answer_<%=Num%>" id="Answer_<%=Num%>"><%=arrQues(5,fnI)%></textarea>
											<input type="hidden" name="QCode_<%=Num%>" id="QCode_<%=Num%>" value="<%=FN_InputVal(arrQues(1,fnI))%>" >
										</td>
									</tr>
									<%
									Next
									%>
								</tbody>
	
							</table>
						</div>
					</form>
					
					<p class="popBtn">
						<a href="javascript:chkQuesForm();"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
						<a href="javascript:;" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
					</p>
				</div>
			</div></div>
			<%
	End Sub

%>
