<%
'____________________________________________________________________________________
'
' * Discription : FnMail.asp / 메일링 함수
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2007-12-21 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________

Sub SB_SendMail(strFrom, strTo, strCC, strSubject, strBody, strFile)

	Const cdoSendUsingMethod = 	"http://schemas.microsoft.com/cdo/configuration/sendusing" 
	Const cdoSMTPServer = 	"http://schemas.microsoft.com/cdo/configuration/smtpserver" 
	Const cdoSMTPServerPort = 	"http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	Const cdoSMTPConnectionTimeout = 	"http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout" 
	Const cdoSMTPAccountName = 	"http://schemas.microsoft.com/cdo/configuration/smtpaccountname" 
	Const cdoSMTPAuthenticate =	"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" 
	Const cdoBasic = 1 	'1 : cdoBasic (기본인증) / 2 : cdoAnoymous (익명엑세스)
	Const cdoSendUserName =	"http://schemas.microsoft.com/cdo/configuration/sendusername" 
	Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword" 
	
	
	Dim objConfig, objMessage, Fields

	
	' ## 메일 발송
	' Get a handle on the config object and it's fields 
	Set objConfig = Server.CreateObject("CDO.Configuration") 
	Set Fields = objConfig.Fields 

	With Fields 
		.Item(cdoSendUsingMethod) = GLOBAL_MAIL_USINGPORT 
		.Item(cdoSMTPServer) = GLOBAL_MAIL_SERVER
		.Item(cdoSMTPServerPort) = 25 
		.Item(cdoSMTPAuthenticate) = cdoBasic
		.Item(cdoSendUserName) = GLOBAL_MAIL_ID
		.Item(cdoSendPassword) = GLOBAL_MAIL_PWD
		.Update 
	End With 

	Set objMessage = Server.CreateObject("CDO.Message") 
	Set objMessage.Configuration = objConfig 

	With objMessage 
	
		.To = strTo
		.From =  strFrom
		.Subject = strSubject
		.HTMLBody = strBody	
		.BodyPart.Charset="ks_c_5601-1987"		' 설정해주지 않을경우 제목 한글 깨짐
		.HTMLBodyPart.Charset="ks_c_5601-1987"

		If strFile <> "" Then
			.AddAttachment = strFile
		End If

		.Send 
	End With 

	Set Fields = Nothing 
	Set objMessage = Nothing 
	Set objConfig = Nothing 
		
End Sub

Function FN_GetMailTemplate(path)

	Dim fnContent : fnContent = ""
	Dim fnObjFs, fnObjFile
	
	Set fnObjFs = Server.CreateObject("Scripting.FileSystemObject")
	Set fnObjFile = fnObjFs.OpenTextFile(Server.MapPath(path),1) 
	
	fnContent = fnObjFile.readall
	fnContent = Replace(fnContent,"[[MAIL_SVR_PATH]]",GLOBAL_DOMAIN)
	fnContent = Replace(fnContent,"[[MAIL_ADM_PATH]]",GLOBAL_DOMAIN & GLOBAL_ADM_PATH)
	
	fnObjFile.Close 

	Set fnObjFile = Nothing 
	Set fnObjFs = Nothing 


	FN_GetMailTemplate = fnContent
End Function
%>