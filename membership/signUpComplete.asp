<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : signUpComplete.asp / HC - 간편회원가입 - 완료 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-24 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Mobile
	Dim yy, mm

	Call Main()
	
	Sub Main()

		Mobile = FN_Req("Mobile","")
	
	End Sub
%>
	<script type="text/javascript">
		function doSignComplete(){
			if ($("#NextUrl").val()){
				window.location.href=$("#NextUrl").val();
			}else{
				window.location.reload();
			}
			
		}
	</script>
	<div class="inner"><div class="inner2">
		<div class="popCont message ">
			<div class="signUpComplete">
				<h2><img src="/images/pages/tit_signUpcomplete.jpg" alt="HC Recruit 의 회원이 되신것을 축하드립니다." /></h2>
				<p class="information">
					회원님의 아이디는 <span><%=Session("UID")%></span>입니다. <br />
					이력서 수정 및 비밀번호 변경은 <a href="/mypage/resume.asp">MY PAGE</a>를 이용해 주세요.
				</p>
				<p class="popBtn">
					<a href="javascript:doSignComplete();" id="btnClose"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				</p>
			</div>
		</div>
	</div></div>
