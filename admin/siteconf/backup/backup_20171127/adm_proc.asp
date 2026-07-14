<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : adm_proc.asp / 관리자정보 저장
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Seq
	Dim AdmID, AdmPwd, AdmName, AdmLevel, AdmTel, AdmMobile, AdmEmail
	Dim strParamFlag
	Dim Result

	Call Main()
	
	Sub Main()
		Call getParam()
		Call setProc()
		Call setResult()
	End Sub
	
	Sub getParam()

		Flag 		= FN_Req("Flag","")
		Seq 		= FN_Req("Seq","0")
	
		AdmID 		= FN_Req("AdmID","")
		AdmPwd 		= FN_Req("AdmPwd","")
		AdmName 	= FN_Req("AdmName","")
		AdmLevel 	= FN_Req("AdmLevel","")
		AdmTel 		= FN_Req("AdmTel","")
		AdmMobile 	= FN_Req("AdmMobile","")
		AdmEmail 	= FN_Req("AdmEmail","")
		
		If Session("ALevel") = "3" And CStr(Session("ASeq")) <> CStr(Seq) Then
			Call SB_ReturnErr("[권한] 잘못된 경로로 접근하셨습니다.","CLOSE")
		End If
		
		If Session("ALevel") = "3" Then
			AdmLevel = Session("ALevel")
		End If
		
		If Flag = "DEL" Then
			strParamFlag = "@StrSeq"
			Seq = Replace(Seq," ","")
		Else
			strParamFlag = "@AdmSeq"
			If Seq = "" Or Not IsNumeric(Seq) Then 	Seq = 0
			If AdmID = "" Or AdmLevel = "" Then 	Call SB_ReturnErr("[인수] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
		If AdmPwd <> "" Then		AdmPwd = hex_sha1(AdmPwd)

		
	End Sub
	
	Sub setProc()
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspAdm_Proc"
			.Parameters.Refresh
			.Parameters("@Flag").Value 			= Flag
			.Parameters(strParamFlag).Value 		= Seq
			.Parameters("@AdmID").Value 		= AdmID
			.Parameters("@AdmPwd").Value 		= AdmPwd
			.Parameters("@AdmName").Value 		= AdmName
			.Parameters("@AdmLevel").Value 		= AdmLevel
			.Parameters("@AdmTel").Value 		= AdmTel
			.Parameters("@AdmMobile").Value 	= AdmMobile
			.Parameters("@AdmEmail").Value 		= AdmEmail
			.Parameters("@RegIP").Value 		= Request.ServerVariables("REMOTE_ADDR")

			.Execute
			
			Result = .Parameters("@Result")
			
		End With
		
	End Sub

	Sub setResult()
		Dim strMsg, strScript
		
		If Result = 0 Then
			strMsg = "처리가 완료됐습니다."
			
			If Flag = "ADD" Or Flag = "MOD" Then
				strScript = "opener.location.reload(); self.close();"
			Else
				strScript = "window.location.href = ""adm_list.asp""; "
			End If
		ElseIf Result = 1 Then
			strMsg = "이미 등록되어 있는 아이디 입니다."
			strScript = "history.back()"
		Else
			strMsg = "처리도중 오류가 발생했습니다."
			strScript = "history.back()"
		End If
	%>
	<!DOCTYPE html>
	<html lang="ko">
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=euc-kr"/>
			<title>관리자 정보 처리 - HC 관리자</title>
			<script type="text/javascript">
			<!--
				alert("<%=strMsg%>");
				<%=strScript%>;
			//-->
			</script>
		</head>
		<body>
		</body>
	</html>
	<%
	End Sub
	
%>
