<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : login.asp / HC - 로그인 폼
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	' 로그인시 접근 제한 페이지
	Call SB_NoLoginPage
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>로그인 - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<style>
	    ::placeholder { /* Chrome, Firefox, Opera, Safari 10.1+ */
			color: #1164c3;
			opacity: 1; /* Firefox */
	    }
		
		:-ms-input-placeholder {
			color: #1164c3 !important;
			font-weight: 400 !important;
		}
	</style>
	<script type="text/javascript">
		var chkFocus = false;
		
		function setPlaceHolderForIE9() {
		    var pos = window.navigator.userAgent.indexOf("MSIE");
			$("input[placeholder]").css("color","#1164c3");
		    if (pos > 0) {
				if (window.navigator.userAgent.substring(pos + 5, window.navigator.userAgent.indexOf(".", pos)) < 10) {
					$("input[placeholder]").each(function () {
						$(this).val($(this).attr("placeholder"));
						$(this).css("color","#1164c3");
					});

					$("input[placeholder]").click(function () {
						if ($(this).val() === $(this).attr("placeholder")) { $(this).val(''); }
						$(this).css("color","#bbbdbf");
					});

					$('input[placeholder]').blur(function () {
					if ($.trim($(this).val()).length === 0) {
						$(this).val($(this).attr("placeholder"));
						$(this).css("color","#1164c3");
					}
					});
				}
		    }
			
			$("input[placeholder]").click(function () {
				if ($(this).val() === $(this).attr("placeholder")) { $(this).val(''); }
				$(this).css("color","#bbbdbf");
			});
			
			$('input[placeholder]').blur(function () {
				if($(this).val() != $(this).attr("placeholder")){
					$(this).css("color","#bbbdbf");
				}else{
					$(this).css("color","#1164c3");	
				}
			});
		}
		
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
					var url = "https://www.mpcjob.co.kr/membership/login_proc.asp"
					param = $("#frmLogin").serialize();
					
					$.ajax({

						type: "POST"
						, url: url
						, data: param
						, crossDomain:true
						, dataType:"jsonp"
						, jsonpCallback:"callback"
						, success: function(content){ 
							if (content.result=="SUCC") {
								if ($("#NextUrl").val()){
									window.location.href=$("#NextUrl").val();
								}else{
									window.location.href="/";
								}
							}else{
								
								if (content.result=="LOCK"){
									alert('로그인 실패 횟수 초과 되었습니다. 담당자에게 연락 바랍니다.');
									window.location.href="/membership/login.asp";
								}
								
								if (content.result=="FAIL"){
								var strHtml = "<h2><img src=\"/images/pages/tit_login2_n01.jpg\" alt=\"입력하신 아이디와 이름이 일치하지 않습니다.\" /></h2>";
								strHtml += "<p class=\"alignC pt10\"><img src=\"/images/pages/txt_login2_n01.jpg\" alt=\"현재 입력하신 아이디가 HC Recruit에 등록되어있지 않거나 아이디 또는 비밀번호를 잘못입력하셨습니다.\" /></p>";
								$("#divLoginMsg").html(strHtml);
								}
							
							}
						}
						, error: function(xhr, status, error) { alert("오류발생"); /*alert("[Rec] " + error);*/}
					});
					return false;
					//rtn = true;
			}
		}

		function txtIDEnter() {
			if (window.event.keyCode == 13 && frmLogin.txtID.value != "") {
				chkFocus = true;
			}
		}
		
		$(function() {
			$(window).load(function(){	
				//$("#txtID").focus();
			});
			$("#txtID").keydown(txtIDEnter);
			$("#txtID").keypress(function(e){ return OnlyNumber_H(e,false); });
			$("#txtID").keyup(function(){ $(this).val( FillNumerics_H($(this).val())); });
			setPlaceHolderForIE9();
		});
	</script>
</head>
<body>
	<div id="wrap" class="login">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div id="divLoginMsg">
					<h2><img src="/images/pages/tit_login.jpg" alt="회원님의 아이디와 비밀번호를 입력해 주세요" /></h2>
				</div>
				
				<form id="frmLogin" name="frmLogin" method="post" onsubmit="javascript: return chkLogin();">
				<input type="hidden" id="NextUrl" name="NextUrl" value="<%=FN_InputVal(Request.ServerVariables("HTTP_REFERER"))%>">
				<input type="hidden" id="PageFlag" name="PageFlag" value="FRONT">
				<div class="loginForm">
					<p class="txt">
						<label for="txtID">아이디</label>
						<input type="text" name="txtID" id="txtID" title="아이디" value="<%=Request.Cookies.Item("UserID")%>" maxlength="15" style="ime-mode:disabled;" placeholder="휴대폰번호 - 포함하여 입력"/>
					</p>
					<p class="txt">
						<label for="txtPwd">비밀번호</label>
						<input type="password" name="txtPwd" id="txtPwd" title="비밀번호" />
					</p>
					<p class="check"><input type="checkbox" name="chkSaveID" id="chkSaveID" title="아이디 저장" value="Y"<%=Fn_SetDefault(Request.Cookies.Item("UserID"),"",""," checked")%> /> <label for="chkSaveID">아이디 저장</label></p>
					<p class="btn"><input type="image" src="/images/pages/btn_login.png" alt="로그인" /></p>
				</div>
				</form>
				<div class="section1">
					<p class="gofind">
						<img src="/images/pages/txt_login_01.jpg" alt="회원님의 아이디와 비밀번호를 잊으셨나요?" /> <a href="find.asp"><img src="/images/pages/btn_find.jpg" alt="아이디/비밀번호 찾기" /></a>
					 </p>
					<p class="gojoin">
						<img src="/images/pages/txt_login_02.jpg" alt="아직 HC Recruit 회원이 아니신가요? 간편가입하세요!" /> <br /><a href="/membership/easySignUp_n.asp"><img src="/images/pages/txt_login_03_n.jpg" alt="간편회원가입하기" /></a>
					</p>
					<p class="note">
						<img src="/images/pages/txt_login_04_n01.gif" alt="입력하신 핸드폰번호가 아이디로 적용됩니다.(-포함 입력 예시: 010-1234-5678)" /><br />
						<img src="/images/pages/txt_login_05_n01.gif" alt="개인정보는 채용정보 제공 시에만 활용됩니다." /><br />
						<img src="/images/pages/txt_login_06_n01.gif" alt="간편회원가입 후 자동 로그인 되어집니다." />
					</p>
				</div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>