<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : SendProc.asp / SMS 발송
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-31 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim receiverList, MessageType, SendMsg, UseMerge, ReservedChk, ResDateTime, callbackPhoneNo
	Dim ResDate, ResHour, ResMin

	Call Main()
	
	Sub Main()
		Call getParam()
		Call getMobileNum()
		Call setResult()
	End Sub
	
	Sub getParam()

		receiverList 	= FN_Req("receiverList","")
		MessageType 	= Trim(FN_Req("MessageType","S"))
		SendMsg 		= Trim(FN_Req("SendMsg",""))
		UseMerge 		= FN_Req("UseMerge","N")
		ReservedChk 	= FN_Req("ReservedChk","0")
		ResDate 		= FN_Req("ResDate",Date())
		ResHour 		= FN_Req("ResHour",Hour(Now()))
		ResMin 			= FN_Req("ResMin",Minute(Now()))
		callbackPhoneNo = FN_Req("callbackPhoneNo",GLOBAL_SMS_CALLBACK)
		
		If FN_isBlank(receiverList) Or FN_isBlank(SendMsg) Then 	Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		
		If ReservedChk = "1" Then 
			ResDateTime = ResDate & " " & ResHour & ":" & ResMin & ":00"
		Else
			ResDateTime = ""
		End If
		
	End Sub
	
	Sub getMobileNum()
		Dim arrRecv : arrRecv = Split(receiverList, Chr(13) + Chr(10))
		Dim i
		
		If IsArray(arrRecv) Then
			' ## 여러명에게 발송
			For i = 0 To UBound(arrRecv)
				Call sendSMS(arrRecv(i))
			Next
		Else
			' ## 한명에게 발송
			Call sendSMS(receiverList)
		End If
	End Sub
	
	Sub sendSMS (argInfo)
		Dim strMobile, strName
		Dim strMsg, strTitle
		
		' 이름과 전화번호 분리		
		If InStr(argInfo,",") > 0 Then
			strMobile = Trim(LEFT(argInfo,Instr(argInfo,",")-1))
			strName = Trim(MID(argInfo,Instr(argInfo,",")+1))
		Else
			strMobile = Trim(argInfo)
			strName = ""	
		End If
		
		' 머지 기능 사용시 처리
		If UseMerge = "Y" Then
			strMsg = Replace(SendMsg,"{이름}",Left(strName,3))
		Else
			strMsg = SendMsg
		End If 
		
		If MessageType = "S" Then
			strTitle = "[MPCJOB] SMS 발송"
		Else
			If InStr(strMsg,".") = 0 Then
				strTitle = Left(strMsg,20) & "..."
			Else
				strTitle = Left(strMsg,InStr(strMsg,"."))
			End If
		End If
		
		Call SB_SendSMS(strMobile, 0, MessageType, strTitle, strMsg, ReservedChk, ResDateTime, "SMS", 0, callbackPhoneNo)
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		Dim strUrl
		%>
		<!DOCTYPE html>
		<html lang="ko">
			<head>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
				<title>SMS 발송 처리 - MPC 관리자</title>
				<script type="text/javascript">
				<!--
					alert("SMS가 등록되었습니다.");
					
					if (window.opener){
						self.close();
					}else{
						window.location.href="/admin/board/mail_list.asp";
					}
				//-->
				</script>
			</head>
			<body>
			</body>
		</html>
		<%
	End Sub
	
%>
