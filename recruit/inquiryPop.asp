<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : inquiryPop.asp / 채용문의 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim RecrSeq
	Dim Mobile, arrMobile
	Dim PrjName, ChargeName, ChargeCode
	

	' 로그인시 접근 제한 페이지
	Call Main()
	
	Sub Main()

		RecrSeq = FN_Req("RecrSeq","0")
		Mobile = Session("UID")
		arrMobile = Split(Mobile,"-")
		
		If UBound(arrMobile) < 2 Then
			ReDim arrMobile(2)
			
			arrMobile(0) = LEFT(Mobile,3)
			arrMobile(1) = MID(Mobile,4, Fn_SetDefault(Len(Mobile),"10","3","4"))
			arrMobile(2) = RIGHT(Mobile,4)
		End If
		
		If RecrSeq = "0" Or Session("UID") = "" Then
			Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","CLOSE_LAYER")
			Response.End
		End If
		
		Call getData
		
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
			.Parameters("@UseHit").Value = 0
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PrjName 	= Rs("PrjName")
				ChargeName 	= Rs("AdmName")
				ChargeCode 	= Rs("ChargeCode")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
%>
<script type="text/javascript">
	function chkForm(){
		
		if (!$("#frmInquiry").find("#SendName").val()){
			alert("이름을 입력해주세요.");
			$("#frmInquiry").find("#SendName").focus();
			return;
		} else if (!$("#frmInquiry").find("#Tel_1").val() || $("#frmInquiry").find("#Tel_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmInquiry").find("#Tel_2").val() || $("#frmInquiry").find("#Tel_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmInquiry").find("#Tel_3").val() || $("#frmInquiry").find("#Tel_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_3").focus();
			return;
		} else if (!$("#frmInquiry").find("#Memo").val()){
			alert("내용을 입력해주세요.");
			$("#frmInquiry").find("#Memo").focus();
			return;
		}
		
		doSubmit();
		
	}
		
	function doSubmit(){
		var url = "inquiryPop_proc.asp"
		param = $("#frmInquiry").serialize();
		
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(content){ 
				if (content=="SUCC") {
					alert("문의하신 내용이 담당자에게 성공적으로 발송 되었습니다.");
					closeLayer();
				}else{
					alert("[채용문의] 문의처리중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}
			}
			, error: function(xhr, status, error) { alert("[채용문의] " + error);}
		});
	}
</script>

	<div class="inner"><div class="inner2">
		<div class="popCont inquiryPop">
			<h2><img src="/images/pages/tit_inquiryPop.jpg" alt="문의하기" /></h2>
			<div class="board_write">
				<h3><%=PrjName%></h3>
				<form name="frmInquiry" id="frmInquiry">
				<input type="hidden" name="RecrSeq" id="RecrSeq" value="<%=RecrSeq%>">
				<input type="hidden" name="ChargeCode" id="ChargeCode" value="<%=ChargeCode%>">
				<input type="hidden" name="ChargeName" id="ChargeName" value="<%=ChargeName%>">
				<table>
					<caption>문의하기 양식</caption>
					<colgroup>
						<col class="w80" />
						<col />
					</colgroup>
					<tbody>
						<tr>
							<th><label for="SendName">이름</label></th>
							<td><input type="text" name="SendName" id="SendName" title="이름" value="<%=Session("UName")%>" class="w150" maxlength="100" /></td>
						</tr>
						<tr>
							<th><label for="Tel_1">휴대폰번호</label></th>
							<td><input type="text" name="Tel_1" id="Tel_1" title="휴대폰번호 앞 3자리" value="<%=arrMobile(0)%>" class="w35" maxlength="3" /> - 
								<input type="text" name="Tel_2" id="Tel_2" title="휴대폰번호 중간 4자리" value="<%=arrMobile(1)%>" class="w40" maxlength="4" /> - 
								<input type="text" name="Tel_3" id="Tel_3" title="휴대폰번호 끝 4자리" value="<%=arrMobile(2)%>" class="w40" maxlength="4" />
							</td>
						</tr>
						<tr>
							<th><label for="Memo">내용</label></th>
							<td>
								<textarea name="Memo" id="Memo" rows="8" cols="40"></textarea>
							</td>
						</tr>
					</tbody>
				</table>
				</form>
			</div>
			<p class="popBtn">
				<a href="javascript:chkForm()"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				<a href="#btnClose" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
			</p>
		</div>
	</div></div>
