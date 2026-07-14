	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<meta name="robots" content="noindex,nofollow"/>
	<link href="/admin/css/admin.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript" src="/common/js/common.js"></script>
	<script type="text/javascript" src="/common/js/common_dev.js"></script>
	<script type="text/javascript" src="/common/js/popup.js"></script>
	<script type="text/javascript">
		function goFrontTest(){
			//alert("테스트 아이디로 서비스모드에 로그인 합니다.");
			//window.open("/membership/login_proc.asp?PageFlag=ADM&txtID=01199283087&txtPwd=1111","","");
			window.open("/","","");
		}
		
		$(function() {
		    $("#btnMyInfo").click( function(){ openWin('/admin/siteconf/adm_form.asp?Seq=<%=Session("ASeq")%>','popAdm',0,0,930,400,0); } );
		});
	</script>

