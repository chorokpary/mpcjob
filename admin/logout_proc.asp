<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : logout_proc.asp / MPCJOB - 관리자 : 로그아웃 처리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	'Session.Abandon	'프론트 동시로그인의 경우 대비 전체 세션 삭제하지 않음
	Session("ASeq") = ""
	Session("AID") = ""
	Session("AName") = ""
	Session("ALevel") = ""	
	
	' OTP 관련 세션 해제 (실서비스 적용 시 아래 주석 해제)
	'Session("TempASeq") = ""
	'Session("TempAID") = ""
	'Session("TempAName") = ""
	'Session("TempALevel") = ""
	'Session("NewOtpSecret") = ""
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>로그아웃 - HC 관리자</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<meta http-equiv="refresh" content="0; url=<%=GLOBAL_ADM_PATH%>">
</head>
<body>
</body>
</html>