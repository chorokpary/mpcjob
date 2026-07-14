<%
'____________________________________________________________________________________
'
' * Discription : FnSMS.asp / SMS 함수
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________

Sub SB_SendSMS(argRecvNum, argRecvSeq, argMsgType, argTitle, argMsg, argIsReserved, argResDate, argPgFlag, argPgSeq, argCallBack)

	Dim RecvNum : RecvNum = Trim(Replace(argRecvNum,"-",""))
	Dim MsgType : MsgType = argMsgType
	Dim StrMsg : StrMsg = argMsg
	
	If MsgType <> "L" Then 	MsgType = "S"
	
	If argCallBack = "" Then 	argCallBack = GLOBAL_SMS_CALLBACK
	
	If MsgType = "S" Then
		Dim StartByte : StartByte = 1
		Dim LoopFlag : LoopFlag = False
		Dim tmpStr : tmpStr = ""
		Dim TotalByte : TotalByte = FN_StringByte(StrMsg)
		
		Do Until LoopFlag
			tmpStr = FN_StringCut(StrMsg, StartByte, StartByte + 90, "")
			StartByte = StartByte + FN_StringByte(tmpStr)

			Call SB_SendSMS_Proc(RecvNum, argRecvSeq, argCallBack, argIsReserved, argResDate, MsgType, argTitle, tmpStr, argPgFlag, argPgSeq)

			If TotalByte < StartByte Then
				LoopFlag = True
			End If
		Loop
		
	Else
		Call SB_SendSMS_Proc(RecvNum, argRecvSeq, argCallBack, argIsReserved, argResDate, MsgType, argTitle, StrMsg, argPgFlag, argPgSeq)
	End If
		
End Sub

Sub SB_SendSMS_Proc (argRecvNum, argRecvSeq, argCallBack, argIsReserved, argResDate, argMsgType, argTitle, argMsg, argPgFlag, argPgSeq)

	'-- 시스템 고정값 ----------------------------------
	CONST registUrl = "http://asp.smsgo.co.kr/Remote/Regist/?"		'-- 메세지 발송 URL
	CONST cancelUrl = "http://asp.smsgo.co.kr/Remote/Modify/?"		'-- 메세지 예약 취소 URL

	Dim url, xmlHttp
	Dim Result, InsResult
	Dim RtnCode : RtnCode = "99|99"
	
	url = registUrl & "serviceAC=" & GLOBAL_SMS_AC & "&sUserID=" & GLOBAL_SMS_ID & "&returnUrl=XML&receiverList=" & argRecvNum &_ 
			"&callbackPhoneNo=" & argCallBack & "&sendMsg=" & argMsg &_
			"&reservedChk=" & argIsReserved & "&resDate=" & argResDate & "&MessageType=" & argMsgType
			'''''''& "&userData1=" & userData1 & "&userData2=" & userData2 & "&userData3=" & userData3

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
		Set RSSItems = xmlDOM.getElementsByTagName("ReturnCode")
		RtnCode = RSSItems(0).Text

		Set RSSItems = xmlDOM.getElementsByTagName("ErrorCode")
		RtnCode = RtnCode & "|" & RSSItems(0).Text
		
		Set xmlDOM = Nothing
		Set RSSItems = Nothing
	
		'Response.ContentType = "text/html"
		'Response.Charset = "euc-kr"
		'Response.Write "<html><body><script> alert('메세지를 발송하였습니다.'); history.go(-1); </script></body></html>"			
	End If

	Set xmlHttp = Nothing
	
	Dim RecrSeq, NoteSeq
	
	RecrSeq = 0
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
		.Parameters("@RtnMsg").Value 		= RtnCode

		'.Parameters(strParamFlag).Value = Seq

		.Execute
			
		InsResult = .Parameters("@Result")
	End With
End Sub
%>