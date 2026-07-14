<%
'____________________________________________________________________________________
'
' * Discription : FnSMS_N.asp / SMS 함수
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


'Call SB_Send_Emfo(strID, "[비밀번호찾기] " & strID & " 회원", "#MPCJOB# 임시 비밀번호를 전송합니다. [" & Passwd & "]")
'Call SB_Send_DB(strID, 0, "S","[비밀번호찾기] " & strID & " 회원", "#MPCJOB# 임시 비밀번호를 전송합니다. [" & Passwd & "]",0,"","MEM",0,Passwd)

Sub SB_Send_Emfo(argRecvNum, argRecvSeq, argMsgType, argTitle, argMsg, argIsReserved, argResDate, argHour, argMin, argPgFlag, argPgSeq, arrUseMerge, arrMerge01, argFromNum)
	
	Dim registUrl, url
	Dim xmlHttp, Result
	Dim RtnCode 

	Dim	m_idx, m_id, m_pwd, call_to, call_from, m_subject, m_message, m_type		
	Dim m_send_type, m_file_name	
	Dim m_yyyy, m_mm, m_dd, m_hh, m_mi		
	Dim url_success, url_fail
	Dim flag_test, flag_deny	
	Dim flag_merge, merge_01, merge_02, merge_03, merge_04, merge_05, return_var	

	Dim i, nameArr, callToArr, tmp_subject, tmp_message

	If UCase(argIsReserved) = "R" Then 
		m_yyyy	= Split(argResDate,"-")(0)
		m_mm	= Split(argResDate,"-")(1)
		m_dd	= Split(argResDate,"-")(2)
	Else
		m_yyyy	= ""
		m_mm	= ""
		m_dd	= ""
	End If 

	If argFromNum = "" Then 
		argFromNum = GLOBAL_SMS_CALLBACK
	End If 

	flag_test = "R" '실제발송 R , 테스트 발송 T
	flag_deny = "Y" '수신거부 Y / N

	registUrl = "http://www.emfohttp.co.kr/send/send_mpc.emfo?"
	url = registUrl & "m_idx=" & make_serial() & "&m_id=" & GLOBAL_SMS_ID_EMFO & "&m_pwd=" & encryption_pw(GLOBAL_SMS_PWD_EMFO) & "&call_to=" & Replace(argRecvNum,",","|") & "&call_from=" & argFromNum & "&m_subject=" & argTitle & "&m_message=" & argMsg & "&m_type=" & argMsgType & "&m_send_type=" & argIsReserved & "&m_file_name=" & m_file_name & "&m_yyyy=" & m_yyyy & "&m_mm=" & m_mm & "&m_dd=" & m_dd & "&m_hh=" & argHour & "&m_mi=" & argMin & "&url_success=" & url_success & "&url_fail=" & url_fail & "&flag_test=" & flag_test & "&flag_deny=" & flag_deny & "&flag_merge=" & arrUseMerge & "&merge_01=" & Replace(arrMerge01,",","|") & "&merge_02=" & Replace(arrMerge01,",","|") & "&merge_03=" & Replace(arrMerge01,",","|") & "&merge_04=" & Replace(arrMerge01,",","|") & "&merge_05=" & Replace(arrMerge01,",","|") & "&return_var=" & return_var 


	Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
	xmlHttp.open "get", url, False
	xmlHttp.send(Request)



	If ( xmlHttp.readyState = 4 And xmlHttp.status = 200 ) Then
		' ### 서버와 정상통신
		Result = xmlHttp.responseText

		
		Dim xmlDOM
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument")
		xmlDOM.async = false
		xmlDOM.LoadXml(Result)
	
		Dim RSSItems
		Set RSSItems = xmlDOM.getElementsByTagName("ResultCode")
		RtnCode = RSSItems(0).Text

		Set RSSItems = xmlDOM.getElementsByTagName("ReturnCode") '한글깨짐으로 영문 / 숫자로 진행 요망
		'RtnCode = RtnCode & "|" & RSSItems(0).Text
		
		Set xmlDOM = Nothing
		Set RSSItems = Nothing
		
	End If

	Set xmlHttp = Nothing


		callToArr	= Split(argRecvNum,"|")
		nameArr		= Split(arrMerge01,"|")
		
		For i = 0 To UBound(callToArr)

				If UCase(arrUseMerge) = "Y" Then 
						tmp_subject = Replace(argTitle,"[조합1]",nameArr(i))
						tmp_message = Replace(argMsg,"[조합1]",nameArr(i))
				Else
						tmp_subject = argTitle
						tmp_message = argMsg
				End If 

				Call SB_Send_DB (callToArr(i), argRecvSeq, tmp_subject, tmp_message, argPgFlag, argPgSeq, RtnCode )
		Next


End Sub 


Sub SB_Send_DB(argRecvNum, argRecvSeq, argTitle, argMsg, argPgFlag, argPgSeq, arrRtnCode)

		Dim RecrSeq, NoteSeq, InsResult
	
		NoteSeq = 0
		
		Select Case argPgFlag
			Case "RECR" : RecrSeq = argPgSeq
			Case "NOTE" : NoteSeq = argPgSeq
		End Select
		
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspSmsHistory_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= "ADD"
			.Parameters("@SendSeq").Value 		= Session("ASeq")
			.Parameters("@RecvSeq").Value 		= argRecvSeq
			.Parameters("@RecvNum").Value 		= argRecvNum
			.Parameters("@Subject").Value 		= argTitle
			.Parameters("@Contents").Value 		= argMsg
			.Parameters("@RecrSeq").Value 		= RecrSeq
			.Parameters("@NoteSeq").Value 		= NoteSeq
			.Parameters("@RtnMsg").Value 		= arrRtnCode

			'.Parameters(strParamFlag).Value = Seq

			.Execute
				
			InsResult = .Parameters("@Result")
		End With

End Sub 
%>