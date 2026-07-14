<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : fine.asp / MPCJOB - 비밀번호 조회
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
	<title>아이디 비밀번호찾기 - MPC JOB</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		var chkFocus = false;

		function chkForm(){
			switch(true) {
				case (!$("#txtID").val()):
					alert("ID를 입력해주세요");
					$("#txtID").focus();
					return false;
					rtn = false;
					break;
				case (!$("#txtName").val() && chkFocus):
					chkFocus = false;
					$("#txtName").focus();
					return false;
					break;
				case (!$("#txtName").val() && !chkFocus):
					alert("이름을 입력해주세요");
					$("#txtName").focus();
					return false;
					break;
				default:
					var url = "find_proc_n.asp"
					param = $("#frmFind").serialize();
					
					$.ajax({
						type: "POST"
						, url: url
						, data: param
						, success: function(content){ 

							if (content=="SUCC") {
								// 성공 레이어 팝업 출력
								openLayer("/membership/find_complete.asp","","w668");
							}else{
								// 실패 레이어 팝업 출력
								openLayer("/membership/find_fail.asp","","w668");
							}
						}
						, error: function(xhr, status, error) { alert("[ERROR] 관리자에게 문의하세요."); /*alert("[Rec] " + error);*/}
					});
					return false;
			}
		}

		function txtIDEnter() {
			if (window.event.keyCode == 13 && frmFind.txtID.value != "") {
				chkFocus = true;
			}
		}
		
		$(function() {
			$(window).load(function(){	
				$("#txtID").focus();
			});
			$("#txtID").keydown(txtIDEnter);
			$("#txtID").keypress(function(e){ return OnlyNumber(e,false); });
			$("#txtID").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		});
	</script>
</head>
<body>
	<div id="wrap" class="find">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<h2><img src="/images/pages/tit_find.jpg" alt="회원님의 아이디와 비밀번호를 잊으셨나요?" /></h2>
				<div class="findForm clearfix">
					<div class="fl_l">
						<h3><img src="/images/pages/stit_find_01.jpg" alt="아이디찾기" /></h3>
						<div class="note">
							<img src="/images/pages/txt_find_01.jpg" alt="회원님의 아이디와 비밀번호를 잊으셨나요?" /><br />
							<img src="/images/pages/txt_find_02.jpg" alt="ID 는 휴대폰 번호 입니다." /><br />
							<img src="/images/pages/txt_find_03.jpg" alt="아이디가 생각나지 않으시거나 휴대폰번호를 변경하신 경우 MPC문의센터 02-3432-6100로 문의 주시기 바랍니다." />
						</div>
					</div>
					<div class="fl_r">
						<h3><img src="/images/pages/stit_find_02.jpg" alt="비밀번호 찾기" /></h3>
						<div class="section2">
						<form id="frmFind" name="frmFind" method="post" onsubmit="javascript: return chkForm();">
							<p class="txt">
								<label for="txtID">아이디</label>
								<input type="text" name="txtID" id="txtID" title="아이디" maxlength="15" style="ime-mode:disabled;"/>
							</p>
							<p class="txt">
								<label for="txtName">이름</label>
								<input type="text" name="txtName" id="txtName" title="이름" maxlength="25" />
							</p>
							<p class="btn"><input type="image" src="/images/pages/btn_confirm.png" alt="확인" /></p>
						</form>
						</div>
					</div>
				</div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->

</body>
</html>