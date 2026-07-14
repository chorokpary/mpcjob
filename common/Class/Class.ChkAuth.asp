<%
'____________________________________________________________________________________
'
' * Discription : Class.ChkAuth.asp / 권한체크
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


Class ClsChkAuth

	Private IsErrDisplay

	'========================================================
	' 클래스 생성자.
	'========================================================
    Sub Class_Initialize()
		IsErrDisplay = True
    End Sub


	'========================================================
	' 클래스 소멸자
	'========================================================
    Sub Class_Terminate
    
    End Sub

	'========================================================
	' 에러정보 출력
	'========================================================	
	Public Sub SetErrDisplay(ByVal as_Boolean)
		SetErrDisplay = Cbool(as_Boolean)
	End Sub

	'========================================================
	' 관리자 체크 : 반환 없음
	'========================================================	
	Public Sub chkAdmin(argMnAuth)
		Dim isAcc : isAcc = False
		Dim AdmLevel, MnLevel

		If IsNumeric(Session("ALevel")) Then	
			AdmLevel = CInt(Session("ALevel"))
		Else
			AdmLevel = 0
		End If

		If IsNumeric(argMnAuth) Then	
			MnLevel = CInt(argMnAuth)
		Else
			MnLevel = 0
		End If
				
		If getIsAdmin() Then
			If AdmLevel > 0 And MnLevel > 0 And  AdmLevel <= MnLevel Then
				isAcc = True
			End If
		End If

		If Not isAcc Then
			' 임시 구문.. 퍼블리셔 작업을 위한.
			If InStr(Request.ServerVariables("PATH_INFO"),"/html/admin/") = 0 Then
		%>
		
		<%
			Response.End
			End If
		End If

	End Sub

	'========================================================
	' Front 권한 체크 : 반환 없음
	'========================================================	
	Public Sub chkLogin(argMnAuth)
		Dim UType, MnLevel
		
		' argMnAuth = 0 : 권한 체크 없음 / 1 : 간편회원 (UType=0) / 2 : 정회원 (UType=1)
		If argMnAuth <> "0" Then
		
			If IsNumeric(Session("UType")) Then	
				UType = CInt(Session("UType"))
			Else
				UType = 0
			End If
	
			If IsNumeric(argMnAuth) Then	
				MnLevel = CInt(argMnAuth)
			Else
				MnLevel = 0
			End If
			
			If Session("UID") = "" And argMnAuth > 0 Then
			' 비회원인데 회원 권한 페이지에 접근 했을 경우
				%>
				<script type="text/javascript">
					window.location.href = "/membership/login.asp";
				</script>
				<%
				Response.End
			ElseIf UType = 0 And argMnAuth = 2 Then
			' 간단 회원인데 정회원 권한 페이지에 접근했을 경우
				%>
				<script type="text/javascript">
					window.location.href = "/mypage/article.asp";
				</script>
				<%
				Response.End
			
			End If
		End If
	End Sub

	'========================================================
	' 관리자 체크
	'========================================================	
	Public Function getIsAdmin()
		If FN_isBlank(Session("ASeq")) Then	
			getIsAdmin = False
		Else
			getIsAdmin = True
		End If
	End Function

	'========================================================
	' 회원 체크
	'========================================================	
	Public Sub chkMember(argUrl)
		Dim Url 	: Url = Request.ServerVariables("URL")
		Dim Param	: Param = Request.ServerVariables("QUERY_STRING") 

		Param = Replace(Param,"&","^")
		
		If FN_isBlank(Session("USeq")) = "" Then
		%>
		<script type="text/javascript">
			<% If InStr(argUrl,"?") > 0 Then %>
			window.location.href = "<%=argUrl%>&ReturnUrl=<%=Url%>?<%=Param%>";
			<% Else %>
			window.location.href = "<%=argUrl%>?ReturnUrl=<%=Url%>?<%=Param%>";
			<% End If %>
		</script>
		<%
			Response.End
		End If
	End Sub

    
End Class
%>