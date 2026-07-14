<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : easySignUp.asp / MPCJOB - 간편회원가입 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim PageFlag, RecrSeq
	Dim yy, mm
	Dim arrCdEzCr, i
	

	' 로그인시 접근 제한 페이지
	Call SB_NoLoginPage()
	Call Main()
	
	Sub Main()

		PageFlag = FN_Req("PageFlag","JOIN")
		RecrSeq = FN_Req("RecrSeq","0")
		
		If PageFlag = "JOIN" Then
			RecrSeq = 0
		End If
		
		Call getCdData
		
	End Sub

	Sub getCdData()
		' 경력코드 DB Get
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "SEL"
			.Parameters("@Tbl").Value 	= "TBL_USER_CAREER_EZ"
			.Parameters("@Field").Value 	= "Career"
			.Parameters("@UseFlag").Value 	= "USE"
			
			' ## 코드 정보
			Set Rs = .Execute


			If Not Rs.eof Then
				arrCdEzCr = Rs.getrows
			End if

			Rs.Close
			Set Rs = Nothing
		End With
	End Sub

%>
<script type="text/javascript">
	function chkJoinForm(){
		
		$("#frmJoin").find("#Flag").val("ADD");

		if (!$("#frmJoin").find("#UsrName").val()){
			alert("이름을 입력해주세요.");
			$("#frmJoin").find("#UsrName").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_Y").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_Y").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_M").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_M").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_D").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_D").focus();
			return;
	    } else if(!IsDate($("#frmJoin").find("#Birthday_Y").val() + "-" + $("#frmJoin").find("#Birthday_M").val() + "-" + $("#frmJoin").find("#Birthday_D").val() )){
			alert("생년월일이 날짜 범위를 벗어납니다.");
	      	$("#frmJoin").find("#Birthday_Y").focus();
			return;			
		} else if (!$("#frmJoin").find("#Mobile_1").val() || $("#frmJoin").find("#Mobile_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_2").val() || $("#frmJoin").find("#Mobile_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_3").val() || $("#frmJoin").find("#Mobile_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_3").focus();
			return;
		} else if (!IsMobile($("#frmJoin").find("#Mobile_1").val() + $("#frmJoin").find("#Mobile_2").val() + $("#frmJoin").find("#Mobile_3").val())){
			alert("핸드폰번호가 바르지 않습니다.");
			$("#frmModify").find("#Mobile_1").focus();
			return;

		} else if (!$("#frmJoin").find("#AddrSido").val()){
			alert("지역을 선택해주세요.");
			$("#frmJoin").find("#AddrSido").focus();
			return;
		} else if (!$("#frmJoin").find("#AddrGugun").val()){
			alert("지역중분류를 선택해주세요.");
			$("#frmJoin").find("#AddrGugun").focus();
			return;
		} else if (!$("#frmJoin").find("#Passwd").val()){
			alert("비밀번호를 입력해주세요.");
			$("#frmJoin").find("#Passwd").focus();
			return;
		}
		
		// 핸드폰번호 중복 체크
		doChkMobile();
		
	}
		
	function doChkMobile() {

		var url = "/common/dev/ajax/check_exists_userid.asp";
		var param = "Mobile=" + $("#Mobile_1").val() + $("#Mobile_2").val() + $("#Mobile_3").val();
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){
					// 가입진행
					doJoinSubmit();
				}else{
					alert("이미 회원가입 되어있으시거나, 다른 사이트를 통해 MPC고객센터에 지원하신 이력이 있습니다. \n\n회원가입 여부를 확인 바랍니다.");
					window.location.href = "/membership/find.asp"
				}
			}
			, error: function(xhr, status, error) {alert("[핸드폰번호 중복 체크] " + error);}
		});
	}

	function doJoinSubmit(){
		var url = "/membership/easySignUp_proc.asp"
		param = $("#frmJoin").serialize();
		
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(content){ 
				if (content=="SUCC") {
					<% If PageFlag = "JOIN" Then %>
					openLayer("/membership/signUpComplete.asp","", "w668");
					<% Else %>
					openLayer("/recruit/apply_proc.asp","RecrSeq=<%=RecrSeq%>&NextAction=RESUM", "w668");
					<% End If %>
				}else if (content=="EXISTS"){
					alert("[회원가입] 중복된 핸드폰번호입니다.");
				}else{
					alert("[회원가입] 가입처리중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}
			}
			, error: function(xhr, status, error) { alert("[회원가입] " + error);}
		});
	}


	function getJoinAddrGugun(){
		var url = "/common/dev/ajax/make_select_gugun.asp";
		var param = "SIDO=" + $(this).val() + "&TargetID=AddrGugun&Css=w115";
		
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(content){ $("#divGugun").html(content) }
			, error: function(xhr, status, error) {alert("[Rec] " + error);}
		});
	}

	$(function() {
	    $("#AddrSido").change(getJoinAddrGugun);
		$("#Mobile_1").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_2").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_3").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_3").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
	});

</script>

	<div class="inner"><div class="inner2">
		<div class="popCont easySignUp">
			<% If PageFlag = "JOIN" Then %>
			<h2><img src="/images/pages/tit_signup.png" alt="간편회원가입" /></h2>
			<p class="note">
				<img src="/images/pages/txt_login_04.jpg" alt="입력하신 핸드폰번호가 아이디로 적용됩니다." /><br />
				<img src="/images/pages/txt_login_05.jpg" alt="개인정보는 채용정보 제공 시에만 활용됩니다." /><br />
				<img src="/images/pages/txt_login_06.jpg" alt="간편회원가입 후 자동 로그인 되어집니다." />
			</p>
			<% Else %>
			<h2><img src="/images/pages/tit_easyRecruit.jpg" alt="간편입사지원" /></h2>
			<p class="note">
				<img src="/images/pages/txt_easyRecruite_01.jpg" alt="간편입사 지원 시, 자동으로 회원가입됩니다.(ID : 휴대폰번호)" /><br />
				<img src="/images/pages/txt_easyRecruite_02.jpg" alt="개인정보는 채용정보 제공 시에만 활용됩니다." /><br />
				<img src="/images/pages/txt_easyRecruite_03.jpg" alt="입사지원 후 자동 로그인 되어집니다." />
			</p>
			<% End If %>
			
			<form id="frmJoin" name="frmJoin" method="post" >
			<input type="hidden" id="Flag" name="Flag">
			<input type="hidden" id="HopePart" name="HopePart" value!="상담원">
			<input type="hidden" id="PageFlag" name="PageFlag" value="<%=PageFlag%>">
			<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
			
			<div class="board_write">
				<table>
					<caption>회원가입양식</caption>
					<colgroup>
						<col class="w80" />
						<col />
						<col class="w80" />
						<col />
					</colgroup>
					<tbody>
						<tr>
							<th><label for="UsrName">이름</label></th>
							<td><input type="text" id="UsrName" name="UsrName" title="이름" value="" class="w195" maxlength="25" /> </td>
							<th><label for="Birthday_Y">생년월일</label></th>
							<td>
								<select name="Birthday_Y" id="Birthday_Y" title="년" class="w55">
									<option value="">년</option>
									<% For yy = Year(Date()) To GLOBAL_SEL_YEAR Step -1  %>
									<option value="<%=yy%>"><%=yy%></option>
									<% Next %>
								</select>
								년
								<select name="Birthday_M" id="Birthday_M" title="월" class="w40">
									<option value="">월</option>
									<% For mm = 1 TO 12 %>
									<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
									<% Next %>
								</select>
								월
								<select name="Birthday_D" id="Birthday_D" title="일" class="w40">
									<option value="">일</option>
									<% For mm = 1 TO 31 %>
									<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
									<% Next %>
								</select>
								일
							</td>
						</tr>
						<tr>
							<th><label for="Mobile_1">휴대폰번호</label></th>
							<td><input type="text" name="Mobile_1" id="Mobile_1" title="휴대폰번호 앞 3자리" class="w35" maxlength="3" style="ime-mode:disabled;" /> - 
								<input type="text" name="Mobile_2" id="Mobile_2" title="휴대폰번호 중간 4자리" class="w35" maxlength="4" style="ime-mode:disabled;" /> - 
								<input type="text" name="Mobile_3" id="Mobile_3" title="휴대폰번호 끝 4자리" class="w35" maxlength="4" style="ime-mode:disabled;" />
							</td>
							<th><label for="Gender">성별</label></th>
							<td>
								<select id="Gender" name="Gender" title="성별선택" class="w65">
									<option value="1">남</option>
									<option value="2">여</option>
								</select>
							</td>
						</tr>
						<tr>
							<th><label for="AddrSido">주소</label></th>
							<td>
								<%
								Dim SidoSql : SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
								response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "AddrSido", "지역", "", "", "", " class=""w85"" ", False)
								%>
								<span id="divGugun">
								<select name="AddrGugun" id="AddrGugun" title="지역중분류" class="w115">
									<option value="">지역중분류</option>
								</select>
								</span>
							</td>
							<th><label for="Passwd">비밀번호</label></th>
							<td><input type="password" name="Passwd" id="Passwd" title="비밀번호" class="w110" maxlength="30" /></td>
						</tr>
					</tbody>
				</table>
			</div>
			
			<% If  Ubound(arrCdEzCr,2) >= 0 Then  %>
			<p class="note2">
				<img src="/images/pages/txt_login_07.jpg" alt="경력이 있으신 곳에 체크 해주세요" />
			</p>
			<div class="check">
				<% For i = 0 To Ubound(arrCdEzCr,2) %>
				<p><input type="checkbox" name="EzCareer" id="EzCareer_<%=i+1%>" value="<%=FN_InputVal(arrCdEzCr(1,i))%>" /><label for="EzCareer_<%=i+1%>"><%=arrCdEzCr(1,i)%></label></p>
				<% Next %>
			</div>
			<% End If %>
			
			</form>
			<p class="popBtn">
				<a href="javascript:chkJoinForm();"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				<a href="#" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
			</p>
		</div>
	</div></div>
