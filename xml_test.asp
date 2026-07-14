<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<!-- #include virtual="/common/Function/FnSms_Md5.asp" -->
<%
	Dim registUrl, url
	Dim xmlHttp, Result
	Dim RtnCode : RtnCode = "99|99"

	Dim	m_idx	
	Dim	m_id		
	Dim	m_pwd		
	Dim	call_to		
	Dim	call_from	
	Dim	m_subject	
	Dim	m_message	
	Dim	m_type		
	Dim	m_send_type	
	Dim	m_file_name	
	Dim	m_yyyy		
	Dim	m_mm		
	Dim	m_dd		
	Dim	m_hh		
	Dim	m_mi		
	Dim	url_success	
	Dim	url_fail	
	Dim	flag_test	
	Dim	flag_deny	
	Dim	flag_merge	
	Dim	merge_01	
	Dim	merge_02	
	Dim	merge_03	
	Dim	merge_04	
	Dim	merge_05	
	Dim	return_var	

	call_to = "01062990919"
	m_subject = "SMS 제목 입니다."
	m_message = "본문 내용 입니다."
	m_type = "SMS"
	m_send_type = "N"
	flag_test = "T"
	flag_deny = "Y"
	flag_merge = "N"
	return_var = "akjdhfkajsdhflkjsadhlkjfhasdlkj"

	'registUrl = "http://www.mpcjob.co.kr/xml.asp?"
	registUrl = "http://www.emfohttp.co.kr/send/send_mpc.emfo?"
	url = registUrl & "m_idx=" & make_serial() & "&m_id=" & GLOBAL_SMS_ID_EMFO & "&m_pwd=" & encryption_pw(GLOBAL_SMS_PWD_EMFO) & "&call_to=" & Replace(call_to,",","|") & "&call_from=" & GLOBAL_SMS_CALLBACK & "&m_subject=" & m_subject & "&m_message=" & m_message & "&m_type=" & m_type & "&m_send_type=" & m_send_type & "&m_file_name=" & m_file_name & "&m_yyyy=" & m_yyyy & "&m_mm=" & m_mm & "&m_dd=" & m_dd & "&m_hh=" & m_hh & "&m_mi=" & m_mi & "&url_success=" & url_success & "&url_fail=" & url_fail & "&flag_test=" & flag_test & "&flag_deny=" & flag_deny & "&flag_merge=" & flag_merge & "&merge_01=" & merge_01 & "&merge_02=" & merge_02 & "&merge_03=" & merge_03 & "&merge_04=" & GLOBAL_SMS_ID & "&merge_05=" & merge_05 & "&return_var=" & return_var 


	Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
	xmlHttp.open "get", url, False
	xmlHttp.send(Request)

	

	If ( xmlHttp.readyState = 4 And xmlHttp.status = 200 ) Then
		' ### 서버와 정상통신
		Result = xmlHttp.responseText
		response.write Result & "<BR>"

		Dim xmlDOM
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument")
		xmlDOM.async = false
		xmlDOM.LoadXml(Result)
	
		Dim RSSItems
		Set RSSItems = xmlDOM.getElementsByTagName("ResultCode")
		RtnCode = RSSItems(0).Text

		Set RSSItems = xmlDOM.getElementsByTagName("ReturnCode") '한글깨짐으로 영문 / 숫자로 진행 요망
		RtnCode = RtnCode & "|" & RSSItems(0).Text
		
		Set xmlDOM = Nothing
		Set RSSItems = Nothing
		
	End If

	Set xmlHttp = Nothing

	Response.write  RtnCode

%>