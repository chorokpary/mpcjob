<!DOCTYPE html>
<html lang="ko">
<%

	username = "testadmin"

	Function GenerateSecret(length)

		chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
		Randomize

		secret = ""

		For i = 1 To length
			secret = secret & Mid(chars, Int(Rnd * Len(chars)) + 1, 1)
		Next

		GenerateSecret = secret

	End Function

	secret = GenerateSecret(16)

	Response.write "secret : " & secret

	qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" & _
			"otpauth://totp/MySite:" & username & "?secret=" & secret & "&issuer=MySite"

%>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<title>HC 관리자</title>
	<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript" src="/common/js/common.js"></script>
	<script type="text/javascript">
	<!--
		var chkFocus = false;
		
		function chkLogin(){
			switch(true) {
				case (!$("#txtID").val()):
					alert("로그인 ID를 입력해주세요");
					$("#txtID").focus();
					return false;
					rtn = false;
					break;
				case (!$("#txtPwd").val() && chkFocus):
					chkFocus = false;
					$("#txtPwd").focus();
					return false;
					break;
				case (!$("#txtPwd").val() && !chkFocus):
					alert("비밀번호를 입력해주세요");
					$("#txtPwd").focus();
					return false;
					break;
				default:
					rtn = true;
			}
		}

		function txtIDEnter() {
			if (window.event.keyCode == 13 && frmLogin.txtID.value != "") {
				chkFocus = true;
			}
		}
		
		
		$(function() {
			$(window).load(function(){	
				$("#txtID").focus();
			});
			$("#txtID").keydown(txtIDEnter);
		});
	//-->
	</script>
</head>

<body>
	<div id="wrap">
		<div class="login">
			<div class="inner">
				<h1><img src="/admin/images/login_admin.jpg" alt="관리자 페이지" /></h1>
				<div class="loginForm">
					<form id="frmLogin" name="frmLogin" action="https://www.mpcjob.co.kr/admin/login_proc.asp" method="post" onsubmit="javascript: return chkLogin();">
					<label for="txtID"><img src="/admin/images/id_admin.jpg" alt="아이디" /></label><input type="text" title="아이디" id="txtID" name="txtID" value="" /><br/>
					<label for="txtPwd"><img src="/admin/images/pw_admin.jpg" alt="비밀번호" /></label><input type="password" title="비밀번호" id="txtPwd" name="txtPwd" value="" />
					<input type="image" src="/admin/images/btn_login.jpg" alt="로그인" class="submit" />
					</form>
				</div>				
			</div>
			<div class="qr-container">
				<img src="<%=qrUrl%>">
			</div>
			<p class="copy"><img src="/admin/images/copy_admin.jpg" alt="COPYRIGHT (C) 2009 MPC LTD., ALL RIGHTS RESERVED" /></p>
		</div>
	</div>
</body>
</html>