<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : logout.asp / HC - 로그아웃 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	'Session.Abandon	'관리자 동시로그인의 경우 대비 전체 세션 삭제하지 않음
	Session("USeq") = ""
	Session("UID") = ""
	Session("UName") = ""
	Session("UType") = ""
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>로그아웃 - HANKOOK CORPORATION</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<meta http-equiv="refresh" content="0; url=/">
</head>
<body>
</body>
</html>