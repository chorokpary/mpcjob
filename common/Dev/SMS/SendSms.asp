<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Function/FnSms_Md5.asp" -->
<%
	Dim strMobile, strTitle, SendMsg

	strMobile 	= FN_Req("strMobile","")
	strTitle 	= FN_Req("strTitle","")
	SendMsg 	= FN_Req("SendMsg","")

	strMobile = "01062990919"
	strTitle = "테스트"
	SendMsg	= "테스트 메세지 입니다"
%>
<!DOCTYPE html>
<html lang="ko">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
		<title>SMS 발송 처리</title>
		<script type="text/javascript">
		//메시지 발송
			function Send_Msg() {

				//발송처리 URL 설정
				var proc_url;
				proc_url = "http://www.emfohttp.co.kr/send/send.emfo";		

				var frm = document.send_form;

				frm.acceptCharset = 'euc-kr';
				if (document.all) document.charset = 'euc-kr';
				
				frm.action = proc_url;
				frm.method = "post";
				frm.submit();
			}

		</script>
	</head>

<body onload="Send_Msg()">

			<form name="send_form" id="send_form" enctype="multipart/form-data" method='post' action="http://www.emfohttp.co.kr/send/send.emfo">

			<input type="hidden" name="m_idx" id="m_idx" value="<%=make_serial()%>" ><br>
			<input type="hidden" name="m_id" id="m_id" maxlength="20" value="<%=GLOBAL_SMS_ID_EMFO%>" ><br>
			<input type="hidden" name="m_pwd" id="m_pwd" maxlength="32" value="<%=encryption_pw(GLOBAL_SMS_PWD_EMFO)%>" ><br>
			<input type="hidden" name="call_to" id="call_to" value="<%=Replace(strMobile,",","|")%>" ><br>
			<input type="hidden" name="call_from" id="call_from" maxlength="12" value="<%=GLOBAL_SMS_CALLBACK%>" ><br>
			<input type="hidden" name="m_subject" id="m_subject" maxlength="40" value="<%=strTitle%>" ><br>
			<input type="hidden" name="m_message" id="m_message" maxlength="2000" value="<%=SendMsg%>" ><br>
			<input type="hidden" name="m_type" id="m_type" maxlength="3" value="SMS" ><br>
			<input type="hidden" name="m_send_type" id="m_send_type" maxlength="1" value="N" ><br>
			<input type="hidden" name="m_file_name" id="m_file_name" maxlength="1" ><br>
			<input type="hidden" name="m_yyyy" id="m_yyyy" maxlength="4" value="" ><br>
			<input type="hidden" name="m_mm" id="m_mm" maxlength="2" value="" ><br>
			<input type="hidden" name="m_dd" id="m_dd" maxlength="2" value="" ><br>
			<input type="hidden" name="m_hh" id="m_hh" maxlength="2" value="" ><br>
			<input type="hidden" name="m_mi" id="m_mi" maxlength="2" value="" ><br>
			<input type="hidden" name="url_success" id="url_success" value="http://www.mpcjob.co.kr/common/Dev/SMS/sendSms_result.asp" ><br>
			<input type="hidden" name="url_fail" id="url_fail" value="http://www.mpcjob.co.kr/common/Dev/SMS/sendSms_result.asp" ><br>
			<input type="hidden" name="flag_test" id="flag_test" maxlength="1" value="T" ><br>
			<input type="hidden" name="flag_deny" id="flag_deny" maxlength="1" value="Y" ><br>
			<input type="hidden" name="flag_merge" id="flag_merge" maxlength="1" value="N" ><br>
			<input type="hidden" name="merge_01" id="merge_01" value="" ><br>
			<input type="hidden" name="merge_02" id="merge_02" value="" ><br>
			<input type="hidden" name="merge_03" id="merge_03" value="" ><br>
			<input type="hidden" name="merge_04" id="merge_04" value="" ><br>
			<input type="hidden" name="merge_05" id="merge_05" value="" ><br>
			<input type="hidden" name="return_var" id="return_var" value="" ><br><br>

			</form>

			</body>
</html>
