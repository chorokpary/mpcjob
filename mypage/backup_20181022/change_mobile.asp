<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : change_mobile.asp / HC - 휴대전화번호 변경 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<script type="text/javascript">
	function chkModForm(){
		var OrgMobile = $("#frmModify").find("#Mobile_1").val() + $("#frmModify").find("#Mobile_2").val() +$("#frmModify").find("#Mobile_3").val();
		var NewMobile = $("#frmModify").find("#New_Mobile_1").val() + $("#frmModify").find("#New_Mobile_2").val() +$("#frmModify").find("#New_Mobile_3").val();
		var ReMobile = $("#frmModify").find("#Re_Mobile_1").val() + $("#frmModify").find("#Re_Mobile_2").val() +$("#frmModify").find("#Re_Mobile_3").val();
		
		if (!$("#frmModify").find("#Mobile_1").val() || $("#frmModify").find("#Mobile_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#Mobile_2").val() || $("#frmModify").find("#Mobile_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmModify").find("#Mobile_3").val() || $("#frmModify").find("#Mobile_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_3").focus();
			return;
		} else if (OrgMobile!="<%=Replace(Session("UID"),"-","")%>"){
			alert("현재 휴대폰번호와 일치하지 않습니다.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_1").val() || $("#frmModify").find("#New_Mobile_1").val().length!=3 ){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_2").val() || $("#frmModify").find("#New_Mobile_2").val().length<3){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_2").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_3").val() || $("#frmModify").find("#New_Mobile_3").val().length!=4){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_3").focus();
			return;
		} else if (OrgMobile == NewMobile){
			alert("기존 번호와 변경할 번호가 동일합니다.");
			$("#frmModify").find("#New_Mobile_1").focus();
			return;
		} else if (!IsMobile($("#frmModify").find("#New_Mobile_1").val() + $("#frmModify").find("#New_Mobile_2").val() + $("#frmModify").find("#New_Mobile_3").val())){
			alert("새로 등록할 핸드폰번호가 바르지 않습니다.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		} else if ($("#ChkMobile").val()=="N"){
			alert("핸드폰번호 중복 체크를 해주세요.");
			return;
		} else if (!$("#frmModify").find("#Re_Mobile_1").val() || $("#frmModify").find("#Re_Mobile_1").val().length!=3 ){
			alert("새로 등록할 핸드폰번호를 다시 입력해주세요.");
			$("#frmModify").find("#Re_Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#Re_Mobile_2").val() || $("#frmModify").find("#Re_Mobile_2").val().length<3){
			alert("새로 등록할 핸드폰번호를 다시 입력해주세요.");
			$("#frmModify").find("Re_Mobile_2").focus();
			return;
		} else if (!$("#frmModify").find("#Re_Mobile_3").val() || $("#frmModify").find("#Re_Mobile_3").val().length!=4){
			alert("새로 등록할 핸드폰번호를 다시 입력해주세요.");
			$("#frmModify").find("#Re_Mobile_3").focus();
			return;
		} else if (NewMobile != ReMobile){
			alert("새 휴대폰번호와 새 휴대폰번호확인 내용이 일치하지 않습니다.");
			$("#frmModify").find("#New_Mobile_1").focus();
			return;
		}
		
		doModSubmit();
		
	}
		
	function doChkMobile() {
		var OrgMobile = $("#frmModify").find("#Mobile_1").val() + $("#frmModify").find("#Mobile_2").val() +$("#frmModify").find("#Mobile_3").val();
		var NewMobile = $("#frmModify").find("#New_Mobile_1").val() + $("#frmModify").find("#New_Mobile_2").val() +$("#frmModify").find("#New_Mobile_3").val();

		if (!$("#frmModify").find("#Mobile_1").val() || $("#frmModify").find("#Mobile_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#Mobile_2").val() || $("#frmModify").find("#Mobile_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmModify").find("#Mobile_3").val() || $("#frmModify").find("#Mobile_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#Mobile_3").focus();
			return;
		} else if (OrgMobile!="<%=Replace(Session("UID"),"-","")%>"){
			alert("현재 휴대폰번호와 일치하지 않습니다.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_1").val() || $("#frmModify").find("#New_Mobile_1").val().length!=3 ){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_1").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_2").val() || $("#frmModify").find("#New_Mobile_2").val().length<3){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_2").focus();
			return;
		} else if (!$("#frmModify").find("#New_Mobile_3").val() || $("#frmModify").find("#New_Mobile_3").val().length!=4){
			alert("새로 등록할 핸드폰번호를 입력해주세요.");
			$("#frmModify").find("#New_Mobile_3").focus();
			return;
		} else if (OrgMobile == NewMobile){
			alert("기존 번호와 변경할 번호가 동일합니다.");
			$("#frmModify").find("#New_Mobile_1").focus();
			return;
		}

		var url = "/common/dev/ajax/check_exists_userid.asp";
		var param = "Mobile=" + NewMobile;
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){
					// 수정진행
					alert("사용 가능한 핸드폰번호입니다.")
					$("#ChkMobile").val("Y");
					doInputActive(1);
				}else{
					// 중복
					alert("사용중인 핸드폰번호입니다.");
					$("#frmModify").find("#New_Mobile_1").val("");
					$("#frmModify").find("#New_Mobile_2").val("");
					$("#frmModify").find("#New_Mobile_3").val("");
				}
			}
			, error: function(xhr, status, error) {alert("[핸드폰번호 중복 체크] " + error);}
		});
	}

	function doModSubmit(){
		var url = "change_proc.asp"
		param = $("#frmModify").serialize();
		
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(content){ 
				if (content=="SUCC") {
					alert("핸드폰번호가 변경되었습니다.");
					window.location.reload();
				}else if (content=="EXISTS"){
					alert("[핸드폰번호 변경] 중복된 핸드폰번호입니다.");
				}else{
					alert("[핸드폰번호 변경] 핸드폰번호 변경중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}
			}
			, error: function(xhr, status, error) { alert("[핸드폰번호 변경] " + error);}
		});
	}
	
	function doInputActive(flag){
		if (flag==1){
			$("#frmModify").find("#Re_Mobile_1").attr("readonly",false);
			$("#frmModify").find("#Re_Mobile_2").attr("readonly",false);
			$("#frmModify").find("#Re_Mobile_3").attr("readonly",false);

			$("#frmModify").find("#Re_Mobile_1").css("background-color","#ffffff");
			$("#frmModify").find("#Re_Mobile_2").css("background-color","#ffffff");
			$("#frmModify").find("#Re_Mobile_3").css("background-color","#ffffff");
		} else {
			$("#frmModify").find("#Re_Mobile_1").attr("readonly",true);
			$("#frmModify").find("#Re_Mobile_2").attr("readonly",true);
			$("#frmModify").find("#Re_Mobile_3").attr("readonly",true);
	
			$("#frmModify").find("#Re_Mobile_1").css("background-color","#eeeeee");
			$("#frmModify").find("#Re_Mobile_2").css("background-color","#eeeeee");
			$("#frmModify").find("#Re_Mobile_3").css("background-color","#eeeeee");
		}
	
	}
	
	$(function() {
		$("#Mobile_1").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_2").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_3").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_3").keyup(function(){ $(this).val( FillNumerics($(this).val())); });

		$("#New_Mobile_1").keypress(function(e){ return OnlyNumber(e,false); });
		$("#New_Mobile_1").keyup(function(){ $(this).val( FillNumerics($(this).val())); $("#ChkMobile").val("N"); doInputActive(0); });
		$("#New_Mobile_2").keypress(function(e){ return OnlyNumber(e,false); });
		$("#New_Mobile_2").keyup(function(){ $(this).val( FillNumerics($(this).val())); $("#ChkMobile").val("N"); doInputActive(0); });
		$("#New_Mobile_3").keypress(function(e){ return OnlyNumber(e,false); });
		$("#New_Mobile_3").keyup(function(){ $(this).val( FillNumerics($(this).val())); $("#ChkMobile").val("N"); doInputActive(0); });

		$("#Re_Mobile_1").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Re_Mobile_1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Re_Mobile_2").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Re_Mobile_2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Re_Mobile_3").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Re_Mobile_3").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
	});

</script>
	<div class="inner"><div class="inner2">
		<div class="popCont easySignUp">
			<h2><img src="/images/pages/tit_phoneMod.jpg" alt="휴대폰번호 변경" /></h2>
			<p class="note">
				<img src="/images/pages/txt_phoneMod_01.jpg" alt="현재 휴대폰 번호를 입력한 후 새로 사용할 휴대폰 번호를 입력하세요." /><br />
				<img src="/images/pages/txt_phoneMod_02.jpg" alt="휴대폰번호 변경 시 아이디가 함께 변경됩니다." />
			</p>
			<div class="board_write">
				<form id="frmModify" name="frmModify" method="post">
				<input type="hidden" id="ChkMobile" name="ChkMobile" value="N">
				<input type="hidden" id="Flag" name="Flag" value="MOBILE_C">
				<table>
					<caption>회원가입양식</caption>
					<colgroup>
						<col class="w140" />
						<col />
					</colgroup>
					<tbody>
						<tr>
							<th><label for="Mobile_1">현재 휴대폰번호</label></th>
							<td><input type="text" name="Mobile_1" id="Mobile_1" class="w35" maxlength="3" />
								- <input type="text" name="Mobile_2" id="Mobile_2" class="w35" maxlength="4" />
								- <input type="text" name="Mobile_3" id="Mobile_3" class="w35" maxlength="4" />
							</td>
						</tr>
						<tr>
							<th><label for="New_Mobile_1">새 휴대폰번호 입력</label></th>
							<td><input type="text" name="New_Mobile_1" id="New_Mobile_1" class="w35" maxlength="3" />
								- <input type="text" name="New_Mobile_2" id="New_Mobile_2" class="w35" maxlength="4" />
								- <input type="text" name="New_Mobile_3" id="New_Mobile_3" class="w35" maxlength="4" />
								<a href="javascript:doChkMobile();"><img src="/images/pages/btn_check.jpg" alt="중복확인" /></a>
							</td>
						</tr>
						<tr>
							<th><label for="Re_Mobile_1">새 휴대폰번호 확인</label></th>
							<td><input type="text" name="Re_Mobile_1" id="Re_Mobile_1" class="w35" maxlength="3" readonly style="background-color:#eeeeee;" />
								- <input type="text" name="Re_Mobile_2" id="Re_Mobile_2" class="w35" maxlength="4" readonly style="background-color:#eeeeee;" />
								- <input type="text" name="Re_Mobile_3" id="Re_Mobile_3" class="w35" maxlength="4" readonly style="background-color:#eeeeee;" />
							</td>
						</tr>
					</tbody>
				</table>
				</form>
			</div>
			<p class="popBtn">
				<a href="javascript:chkModForm();"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				<a href="#" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
			</p>
		</div>
	</div></div>
