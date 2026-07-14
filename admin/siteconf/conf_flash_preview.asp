<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_flash_preview.asp / 플래시 미리보기
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-08-13 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>플래시 관리 미리보기 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body class="pop" style="width:641px;">
	<div id="wrap" style="min-width:0;">
		<div class="btnArea">
			<div class="fl_r">
				<span class="btns btn_g"><a href="#Close" id="btnClose">x 닫기</a></span>
			</div>
		</div>
		<div id!="contArea">
			<script type="text/javascript"> getFlash("/images/flash/main.swf", 641, 341); </script>
		</div>

	</div>
</body>
</html>