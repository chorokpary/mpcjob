<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_skin_list.asp / 전체 채용공고 목록 스킨
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2014-03-17 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>전체 채용공고 -HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>전체 채용공고</h2>
					<%
					Dim CONF_ISVIEW : CONF_ISVIEW = "-1"
					%>
					<!--#include file="job_list.asp"-->
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>