<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : sms.asp / 관리자 SMS 발송 팝업
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-31 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>결과통보 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body class="pop">
	<div id="wrap" class="sms">
		<div id="contArea">
			<h2>결과통보</h2>
			<!--#include virtual="/common/Dev/SMS/SendForm_n.asp"-->
			<div class="btnArea">
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>