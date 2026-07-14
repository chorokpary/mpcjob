<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pop_message.asp / 팝업 쪽지 알림
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-08-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
Dim NoteSeq : NoteSeq = FN_Req("Seq","0")
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function goNote(){
			opener.location.href="/admin/message/message_view.asp?Seq=<%=NoteSeq%>";
			self.close();
		}
	//-->
	</script>
</head>

<body class="popMessage">
	<div id="wrap" class="">
		<img src="/admin/images/pop_message.jpg" alt="" border="0" usemap="#popMessage" />
		<map name="popMessage">
			<area shape="rect" coords="156,242,222,276" href="javascript:goNote();">
			<area shape="rect" coords="226,242,294,278" href="javascript:self.close();">
		</map>
	</div>
</body>
</html>