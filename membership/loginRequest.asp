<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : loginRequest.asp / HC - 로그인요청 메시지 (Ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<% 
	Dim RecrSeq

	Call Main()
	
	Sub Main()
		
		Call getParam()
		
	End Sub

	Sub getParam()
	
		RecrSeq 	= FN_Req("RecrSeq","0")
		
	End Sub
	

%>
<script type="text/javascript">
	function goLogin(){
		$("#frmPage").attr({action:"/membership/login.asp", method:"get"}).submit();
	}
</script>
	<div class="inner"><div class="inner2">
		<div class="popCont message">
			<div class="loginRequest">
				<h2><img src="/images/pages/tit_loginRequest.jpg" alt="로그인이 필요한 페이지 입니다." /></h2>
				<p class="information">
					<img src="/images/pages/txt_loginRequest.jpg" alt="비회원이시면 간편입사지원을 해주세요." />
				</p>
				<p class="popBtn">
					<a href="javascript:goLogin();"><img src="/images/pages/btn_request_login.jpg" alt="로그인" /></a><a href="javascript:doJoin('RECR','<%=RecrSeq%>');"><img src="/images/pages/btn_request_support.jpg" alt="간편입사지원" /></a>
				</p>
			</div>
		</div>
	</div></div>
