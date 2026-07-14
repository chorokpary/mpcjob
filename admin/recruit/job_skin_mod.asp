<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_skin_mod.asp / 수정페이지 스킨
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>채용공고 수정 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea" class="job_reg">
					<%
					Dim CONF_FLAG : CONF_FLAG = FN_Req("Flag","MOD")
					%>
					<!--#include file="job_form.asp"-->
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>