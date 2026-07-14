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

	Dim strMobile, strName
	Dim strMsg, strTitle

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setResult()
	End Sub
	
	Sub getParam()

		receiverList 	= unescape(FN_Req("receiverList",""))
		MessageType 	= Trim(FN_Req("MessageType","S"))
		SendMsg 		= unescape(Trim(FN_Req("SendMsg","")))
		UseMerge 		= FN_Req("UseMerge","N")
		ReservedChk 	= FN_Req("ReservedChk","N")
		ResDate 		= FN_Req("ResDate",Date())
		ResHour 		= FN_Req("ResHour",Hour(Now()))
		ResMin 			= FN_Req("ResMin",Minute(Now()))
		callbackPhoneNo = FN_Req("callbackPhoneNo",GLOBAL_SMS_CALLBACK)

		Dim arrRecv : arrRecv = Split(receiverList, Chr(13) + Chr(10))
		Dim i
		
		If IsArray(arrRecv) Then

			' ## 여러명에게 발송
			For i = 0 To UBound(arrRecv)
				If InStr(arrRecv(i),",") > 0 Then
					strMobile = strMobile &  Trim(LEFT(arrRecv(i),Instr(arrRecv(i),",")-1)) & "|"
					strName = strName & Trim(MID(arrRecv(i),Instr(arrRecv(i),",")+1)) & "|"
				Else
					strMobile = strMobile & Trim(arrRecv(i)) & "|"
					strName = strName & "" & "|"
				End If
			Next

			strMobile =  Trim(LEFT(strMobile,Len(strMobile)-1))
			strName =  Trim(LEFT(strName,Len(strName)-1))
		Else
			' ## 한명에게 발송
			strMobile = Trim(LEFT(arrRecv(i),Instr(arrRecv(i),",")-1)) 
			strName = ""
		End If


		If UCase(MessageType) = "L" Then 	
			MessageType = "LMS"

			If InStr(SendMsg,".") = 0 Then
				strTitle = Left(SendMsg,37) & "..."
			Else
				strTitle = Left(SendMsg,InStr(SendMsg,"."))
			End If
		Else
			MessageType = "SMS"
			strTitle = "[HC] SMS 발송"
		End If 

	
		' 머지 기능 사용시 처리
		If UseMerge = "Y" Then
			strTitle = Replace(strTitle,"{이름}","[조합1]")
			SendMsg = Replace(SendMsg,"{이름}","[조합1]")
		End If 

		If Len(strTitle) > 40 Then 
			strTitle = Left(SendMsg,37) & "..."
		End If 

		If FN_isBlank(receiverList) Or FN_isBlank(SendMsg) Then 	Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		

		Call SB_Send_Emfo(strMobile, 0, MessageType, strTitle, SendMsg, ReservedChk, ResDate, ResHour, ResMin, 0, 0, UseMerge, strName, callbackPhoneNo)

	End Sub
	


	Sub setResult()
		Dim strScript
		Dim strUrl

		%>
		<!DOCTYPE html>
		<html lang="ko">
			<head>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
				<title>SMS 발송 처리 - HC 관리자</title>
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
			<body >
			</body>
		</html>
		<%
	End Sub
	
%>
