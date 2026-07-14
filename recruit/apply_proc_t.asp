<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : apply_proc.asp / 입사지원 처리 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-25 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim RecrSeq, NextAction, PassQues, DtSize

	Dim arrQues, arrQuesLen, Title
	
	Dim Result

	Dim PartnerDepth1, PartnerDepth2, RecrStatus
	Dim Recommend_num, Recommend_name, Recommend_center
	
	Call Main()
	
	Sub Main()
	
		RecrSeq 	= FN_Req("RecrSeq","0")
		NextAction = FN_Req("NextAction","")
		PassQues = FN_Req("PassQues","N")	' 사전질문 답변 여부 체크
		DtSize = FN_Req("DtSize","0")

		
		PartnerDepth1 = FN_Req("PartnerDepth1","")
		PartnerDepth2 = FN_Req("PartnerDepth2","")

		If PartnerDepth1 = "직원추천" Then 
			PartnerDepth1 = "내부추천"
		End If 

		Recommend_num = FN_Req("recommend_num","")
		Recommend_name = FN_Req("recommend_name","0")
		Recommend_center = FN_Req("recommend_center","0")
		

		If PassQues <> "Y" Then	PassQues = "N"

		If Session("USeq") = "" Then 
		%>
			<script type="text/javascript">
				alert("입사지원 처리도중 오류가 발생했습니다.\n로그인 후 다시 시도하세요.");
				location.href = "/membership/login.asp"
			</script>
		<%
			Response.End 
		End If 

			
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
				' 입사지원
				Call setApplyProc()
			End If
		Else
			Result = 99
		End If
			
		Call getData
		Call setApplyResult()

	End Sub


	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= RecrSeq
			.Parameters("@UseHit").Value = 1
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				Title 		= FN_InputVal(Rs("Title"))
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub


	' #### 입사지원
	Sub setApplyProc()
		Dim Rs
		Dim tmpResult : tmpResult = 0
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_Proc_2"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "ADD_USR_R"
			.Parameters("@RecrSeq").Value 		= RecrSeq
			.Parameters("@UsrSeq").Value 		= Session("USeq")
			
			.Parameters("@PartnerDepth1").Value = PartnerDepth1
			.Parameters("@PartnerDepth2").Value = PartnerDepth2

			.Parameters("@Recommend_Num").Value = Recommend_num
			.Parameters("@Recommend_Name").Value = Recommend_name
			.Parameters("@Recommend_Center").Value = Recommend_center

			.Parameters("@Status").Value = RecrStatus
			
			.Execute
			
			tmpResult = .Parameters("@Result")
			
		End With
		
		If tmpResult <> 0 Then
			Result = 90
		Else
			Result = 0
		End If

	End Sub

	Sub setApplyResult()
		Dim strMsg, strScript
		Dim strUrl
		
		If Result = 0 Then
			If NextAction = "RESUM" Then
			%>
			<script type="text/javascript">
				alert("입사지원이 완료되었습니다. 이력서 관리 페이지로 이동합니다.");
				window.location.href="/mypage/resume.asp";
			</script>
			<%
			Else
			%>
			<div class="inner"><div class="inner2">
				<div class="popCont message ">
					<div class="recruitComplete">
						<h1><%=Title%></h1>
						<h2><img src="/images/pages/tit_recruitComplete.jpg" alt="입사지원이 완료되었습니다." /></h2>
						<p class="information">
							<img src="/images/pages/txt_recruitComplete.jpg" alt="입사지원내역을 확인 하실 수 있는 입사지원 관리 페이지로 이동하시겠습니까?" />
						</p>
						<p class="popBtn">
						<a href="/mypage/apply.asp"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
						<a href="javascript:;" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
					</p>
					</div>
				</div>
			</div></div>
			<%
			End If
		Else 
			%>
			<script type="text/javascript">
				alert("입사지원 처리도중 오류가 발생했습니다. 다시 시도하세요.");
				window.location.reload();
			</script>
			<%
		End If
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
				arrQuesLen = -1
				'사전 질문이 없을경우 바로 지원 프로세스 타도록..
				PassQues = "Y"
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
					
					if (!len){
						alert("잘못된 경로로 접근하셨습니다.");
						$("#btnClose").click();
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
						var url = "/recruit/apply_proc_t.asp"
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
						<input type="hidden" id="NextAction" name="NextAction" value="<%=NextAction%>">
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
