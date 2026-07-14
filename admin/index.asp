<!-- #include virtual="/common/CommonConfig.asp" -->
<!DOCTYPE html>
<html lang="ko">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<%
		Dim clsAuth		: Set clsAuth = New ClsChkAuth
		
		'로그인 체크
		'clsAuth.chkAdmin()
		If clsAuth.getIsAdmin() Then
			%><meta http-equiv="refresh" content="0; url=recruit/job_skin_list.asp"><%
		Else
			%><meta http-equiv="refresh" content="0; url=login.asp"><%
		End If
		
		Set clsAuth = Nothing
	
	%>
	<title>HC 관리자</title>
	<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript" src="/common/js/common.js"></script>
</head>
<body>
</body>
</html>
