<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : change_passwd.asp / HC - 비밀번호 변경 (Layer Popup)
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
		
		if (!$("#frmModify").find("#Passwd").val()){
			alert("현재 비밀번호를 입력해주세요.");
			$("#frmModify").find("#Passwd").focus();
			return;
		} else if (!$("#frmModify").find("#New_Passwd").val()){
			alert("새 비밀번호를 입력해주세요.");
			$("#frmModify").find("#New_Passwd").focus();
			return;
		} else if (!$("#frmModify").find("#Re_Passwd").val()){
			alert("새 비밀번호를 다시 입력해주세요.");
			$("#frmModify").find("#Re_Passwd").focus();
			return;
		} else if ($("#frmModify").find("#New_Passwd").val() != $("#frmModify").find("#Re_Passwd").val()){
			alert("입력하신 새 비밀번호와 새 비밀번호 확인이 동일하지 않습니다.");
			$("#frmModify").find("#New_Passwd").focus();
			return;
		}
		
		doChkPasswd();
	}
		
	function doChkPasswd() {
		var url = "/common/dev/ajax/check_user_passwd.asp";
		var param = "Passwd=" + $("#frmModify").find("#Passwd").val();
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){
					// 변경진행
					doModSubmit();
				}else{
					// 비밀번호 불일치
					alert("입력하신 비밀번호가 맞지 않습니다.");
				}
			}
			, error: function(xhr, status, error) {alert("[비밀번호 체크] " + error);}
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
					alert("비밀번호 변경이 완료 되었습니다.");
					window.location.reload();
				}else{
					alert("[비밀번호 변경] 비밀번호 변경중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}
			}
			, error: function(xhr, status, error) { alert("[비밀번호] " + error);}
		});
	}
	
</script>
	<div class="inner"><div class="inner2">
		<div class="popCont easySignUp">
			<h2><img src="/images/pages/tit_passMod.jpg" alt="비밀번호 변경" /></h2>
			<p class="note">
				<img src="/images/pages/txt_passMod.jpg" alt="현재 비밀번호를 입력한 후 새로 사용할 비밀번호를 입력하세요." />
			</p>
			<div class="board_write">
				<form id="frmModify" name="frmModify" method="post">
				<input type="hidden" id="Flag" name="Flag" value="PWD_C">
				<table>
					<caption>회원가입양식</caption>
					<colgroup>
						<col class="w135" />
						<col />
					</colgroup>
					<tbody>
						<tr>
							<th><label for="">현재 비밀번호</label></th>
							<td><input type="password" name="Passwd" id="Passwd" class="w130" maxlength="30" /></td>
						</tr>
						<tr>
							<th><label for="">새 비밀번호 입력</label></th>
							<td><input type="password" name="New_Passwd" id="New_Passwd" class="w130" maxlength="30" /></td>
						</tr>
						<tr>
							<th><label for="">새 비밀번호 확인</label></th>
							<td><input type="password" name="Re_Passwd" id="Re_Passwd" class="w130" maxlength="30" /></td>
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
