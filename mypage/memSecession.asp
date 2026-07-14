<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
	Call SB_LoginChkPage()
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>MY PAGE - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function chkForm(){
			if (!$("#frmInfo").find("#UserID").val()){
				alert("아이디를 입력해주세요.");
				$("#frmInfo").find("#UserID").focus();
				return;
			} else if (!$("#frmInfo").find("#UserName").val()){
				alert("이름을 입력해주세요.");
				$("#frmInfo").find("#UserName").focus();
				return;
			} else if (!$("#frmInfo").find("#Passwd").val()){
				alert("비밀번호를 입력해주세요.");
				$("#frmInfo").find("#Passwd").focus();
				return;
			} else if (!$("#frmInfo").find("#RePasswd").val()){
				alert("비밀번호확인을 입력해주세요.");
				$("#frmInfo").find("#RePasswd").focus();
				return;
			} else if ($("#frmInfo").find("#Passwd").val() != $("#frmInfo").find("#RePasswd").val()){
				alert("비밀번호가 확인란과 일치하지 않습니다.");
				$("#frmInfo").find("#RePasswd").focus();
				return;
			}

			$("#frmInfo").attr({action:"https://www.mpcjob.co.kr/mypage/memSecession_proc.asp", method:"POST"}).submit();
		}
		
		function doReset(){
			 $("#frmInfo")[0].reset();
		}
	</script>
</head>
<body>
	<div id="wrap" class="memSecession">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_memSecession.jpg" alt="회원탈퇴" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="section">
					<h3><img src="/images/pages/stit_memSecession.jpg" alt="항상 열린공간으로서의 HC Recruit가 되겠습니다." /></h3>
					<p><img src="/images/pages/txt_memSecession.jpg" alt="회원탈퇴를 하시려면 사용하시는 아이디와 이름, 비밀번호를 정확하게 입력하시기 바랍니다." /></p>
					
					<form id="frmInfo" name="frmInfo" method="POST">
					<div class="board_write">
						<table>
							<caption>회원가입양식</caption>
							<colgroup>
								<col class="w136" />
								<col />
							</colgroup>
							<tbody>
								<tr>
									<th><label for="">아이디</label></th>
									<td><input type="text" name="UserID" id="UserID" title="아이디" class="w165" maxlength="15" style="ime-mode:disabled;" /></td>
								</tr>
								<tr>
									<th><label for="">이름</label></th>
									<td><input type="text" name="UserName" id="UserName" title="이름" class="w165" maxlength="25" /></td>
								</tr>
								<tr>
									<th><label for="">비밀번호</label></th>
									<td><input type="password" name="Passwd" id="Passwd" title="비밀번호" class="w165" maxlength="30" /></td>
								</tr>
								<tr>
									<th><label for="">비밀번호확인</label></th>
									<td><input type="password" name="RePasswd" id="RePasswd" title="비밀번호확인" class="w165" maxlength="30" /></td>
								</tr>
							</tbody>
						</table>
					</div>
					</form>
					
					<p>
						<a href="javascript:chkForm()"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
						<a href="javascript:doReset();"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
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