<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : apply_confirm.asp / 입사지원 취소 확인 창 (ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq : Seq = FN_Req("Seq","0")
%>
	<div class="inner"><div class="inner2">
		<div class="popCont applyMng ">
			<div class="recruitComplete">
				<h2><img src="/images/pages/tit_applyMngCancel.jpg" alt="입사지원을 취소하시겠습니까?" /></h2>
				<p class="popBtn">
				<a href="javascript:doDel(<%=Seq%>);"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				<a href="javascript:;" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
			</p>
			</div>
		</div>
	</div></div>
