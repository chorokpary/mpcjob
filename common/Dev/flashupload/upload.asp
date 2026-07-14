<%
	
	Dim tmpDt
	tmpDt = NOW()
	Session("DirPath") = Year(tmpDt) & "/" &_
			Year(tmpDt) &_
			Right("00" & Month(tmpDt),2) &_
			Right("00" & Day(tmpDt),2) &_
			Right("00" & Hour(tmpDt),2) &_
			Right("00" & Minute(tmpDt),2) &_
			Right("00" & Second(tmpDt),2) &_
			"_" & Session.SessionID
%>
<html>
	<head>
		<meta http-equiv="Content-type" content="text/html; charset=utf-8">
		<meta http-equiv="cache-control" content="no-cache, must-revalidate">
		<meta http-equiv="pragma" content="no-cache">
		<title>이미지삽입</title>
		<link rel="stylesheet" type="text/css" href="/common/Dev/SmartEditor/css/editor.css" />
		<link href="upload.css" rel="stylesheet" type="text/css" />
		<script type="text/javascript" src="/common/Js/jquery-1.4.2.min.js"></script>
		<script type="text/javascript" src="flash.js"></script>
		<script language="javascript">
		
			function closeWin() {
			  parent.parent.oEditors.getById["<%=request("id")%>"].exec("SE_TOGGLE_IMAGEUPLOAD_LAYER");
			  return false;
			}
			
			// ### 파일 업로드 START
			/* 제한 사이즈 초과시 에러 메시지 출력 */
			function overSize( limitSize ) {
				var message = "허용 업로드 사이즈는 " + limitSize + "입니다.";
				alert( message );
			}
		
			/* 플래쉬 파일 리스트에 추가된 파일을 전송 : ie에서만 실행 */
			function fileUpload() {
				//var movie = document.getElementById("upload");
				var movie = window.document.upload;
				var totalSize = movie.getVariable("totalSize");
		
				// 업로드된 파일이 있을 경우 실행
				if( totalSize > 0 ) {
					movie.SetVariable( "fileUpload", "" );
					// 전송 버튼 비활성화
					//form.send_button.disabled = true;
				} else {
					alert("업로드할 파일을 선택해 주세요.");
				}
			}
		
			/* 파일 업로드 완료 후 처리 메쏘드 */
			function send()	{
				
				//var objThumb;
				//objThumb = parent.parent.document.getElementById("thumb");
				getFileList();
				
				closeWin();
			}
		
			/* 허용하지 않은 형식의 파일일 경우 에러 메시지 출력 */
			function fileTypeError( notAllowFileType )	{
				alert("확장자가 " + notAllowFileType + "인 파일들은 업로드 할수 없습니다.");
			}
		
			/* 플래쉬 버전이 8 미만일 경우 에러 메시지 출력 */
			function versionError() {
				alert("플래쉬 버전을 확인하여 주십시오.( 버전 8 이상만 가능합니다. )\n이미 설치하신 경우는 브라우저를 전부 닫고 다시 시작하여 주십시오.");
			}
		
			/* http 에러 발생시 출력 메시지 */
			function httpError(str) {
				//alert(str);
				alert("네트워크 에러가 발생하였습니다. 잠시후 다시 시도하여 주십시오.");
			}
		
			/* 파일 입출력 에러 발생시 출력 메시지 ( 주로 업로드 하려는 파일을 다른 프로그램에서 사용중일 경우 발생 ) */
			function ioError() {
				alert("입출력 에러가 발생하였습니다.\n( 다른 프로그램에서 이 파일을 사용하고 있는지 확인하신 후 다시 시도하여 주십시오. )");
			}

			/* 파일 업로드 후 목록 가져오기 */
			function getFileList() {
				var url = "./upload_fileslist.asp";
				var pars = "";

				$.ajax({
					type: "POST"
					, dataType: "xml"
					, url: url
					, data: pars
					, success: getFIleListResponse    //function result 를 의미함
					, error: function(xhr, status, error) {alert(error); }
				});
	        }
	
			function getFIleListResponse(RData) {
				var date_len = $(RData).find("url").length;
				if (date_len > 0){
					$(RData).find("url").each(function(idx) {              
						parent.parent.insertIMG('<%=request("id")%>',$(this).text());
					});
				}
			}

		</script>
	</head>
	<body>
		<div id="naver_common_editor">
			<form id="editor_upimage" name="editor_upimage" method="post">
			<input type="hidden" name="id" value="<%=request("id")%>">
			<div class="pic_content" style="border:0;">
				<div style="position:absolute; top:20px;left:110px;;text-align:right;font-size:8pt;padding-left:30px; color:#FF8000; padding-right:13px;">이미지명은 영문만</div>
				<script type="text/javascript">
				<!--
					flash(400,120,"/common/Dev/flashupload/upload.swf","upload");
				//-->
				</script>
				<div class="btn_box">
					<img src="img/btn_layer_confirm.gif" alt="확인" width="38" height="21" onclick="fileUpload()" style="cursor:pointer">
					<img src="img/btn_layer_cancel.gif" alt="취소" width="38" height="21" border="0" onclick="closeWin()" style="cursor:pointer">
				</div>
			</div>
			</form>
		</div>
	</body>
</html>
