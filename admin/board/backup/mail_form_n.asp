<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mail_form.asp / SMS 발송 폼
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-30 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>메일링 서비스 - MPC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body>
	<div id="wrap" class="sms">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>Mailling 서비스 - SMS 발송</h2>
					<!--#include virtual="/common/Dev/SMS/SendForm_n.asp"-->
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>