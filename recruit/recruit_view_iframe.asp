<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : recruit_view_ifame.asp / 채용공고 상세보기 iframe
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-08-03 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim txtContents : txtContents = FN_Req("txtContents","")
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>채용정보 - HANKOOK CORPORATION</title>
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript">
		// 미리보기 > 로컬이미지
		$(window).load(function(){
			$("#ifrmContents", window.parent.document).css("height",$("#divWrap").height() + 20); 
		})
	</script>
	<style type="text/css" media="screen">
		#divWrap{font-size:12px;}
	</style>
</head>
<body>
	<div id="divWrap">
	<%=txtContents%>
	</div>
</body>
</html>