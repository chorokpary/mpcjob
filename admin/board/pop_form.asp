<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pop_form.asp / 팝업 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-22 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim Title, IsView, Width, Height, BDate, EDate, Contents
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
			IsView = True
			
			Width = "400"
			Height = "300"
			BDate = Date()
			EDate = Date()
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfPopup_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				Title 	= FN_InputVal(Rs("Title"))
				IsView 	= Rs("IsView")
				Width 	= Rs("Width")
				Height 	= Rs("Height")
				BDate 	= Rs("BDate")
				EDate 	= Rs("EDate")
				Contents = Rs("Contents")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>팝업관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript" src="/common/Dev/SmartEditor/js/HuskyEZCreator.js" charset="utf-8"></script>
	<script type="text/javascript">
	<!--
		var oEditors = [];

		function chkForm(){
			if (!$("#Title").val()){
				alert("제목을 입력해 해주세요.");
				$("#Title").focus();
				return;
			}
			if (!$("#Width").val()){
				alert("가로사이즈를 입력해주세요.");
				$("#Width").focus();
				return;
			}
			if (!IsNumerics($("#Width").val())){
				alert("가로사이즈는 숫자만 입력해주세요.");
				$("#Width").focus();
				return;
			}
			if (!$("#Height").val()){
				alert("세로사이즈를 입력해주세요.");
				$("#Height").focus();
				return;
			}
			if (!IsNumerics($("#Height").val())){
				alert("세로사이즈는 숫자만 입력해주세요.");
				$("#Height").focus();
				return;
			}
			if ($("#BDate").val() > $("#EDate").val()){
				alert("게시 시작일이 등록일보다 느린 날짜입니다.");
				return;
			}
		
			oEditors.getById["Contents"].exec("UPDATE_CONTENTS_FIELD", []);
			
			$("#frmInfo").attr({action:"pop_proc.asp", method:"post"}).submit();
		}
		
		function goPreview(){
			oEditors.getById["Contents"].exec("UPDATE_CONTENTS_FIELD", []);
			openWin("about:blank","popPv",0,0,$("#Width").val(),$("#Height").val(),0);
			$("#frmInfo").attr({action:"pop_preview.asp", method:"post", target:"popPv"}).submit();
		}
		
		function doReset(){
			$("#frmInfo")[0].reset();

			var content = $("#Contents_org").val();
			oEditors.getById["Contents"].exec("SET_IR", [content]);
		}

		
		$(document).ready(function() {
			// ### DatePicker
			$("#BDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#EDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});
			
			// ### WYSIWYG
			nhn.husky.EZCreator.createInIFrame({
				oAppRef: oEditors,
				elPlaceHolder: "Contents",
				sSkinURI: "/common/Dev/SmartEditor/SmartEditor2Skin.html",	
				fCreator: "createSEditor2"
			});
			
			// ### Button Control
			$("#btnCancel").click(doReset);
			$("#btnPreview").click(goPreview);
			$("#btnSubmit").click(chkForm);
		});
	//-->
	</script>
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>팝업관리</h2>

					<div class="tblArea">
						<form id="frmInfo" name="frmInfo" action="adm_action.asp" method="post">
						<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
						<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
						<table class="tblForm">
							<colgroup>
								<col class="w85" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="Title">제목</label></th>
									<td>
										<input type="text" id="Title" name="Title" title="제목" value="<%=Title%>" maxlength="100" class="w660" /> 
									</td>
								</tr>
								<tr>
									<th><label for="IsView_Y">공개옵션</label></th>
									<td>
										<input type="radio" name="IsView" value="1" id="IsView_Y"<%=Fn_SetDefault(IsView,"True"," checked","")%>/><label for="">공개</label>
										<input type="radio" name="IsView" value="0" id="IsView_N"<%=Fn_SetDefault(IsView,"False"," checked","")%>/><label for="">비공개</label>
									</td>
								</tr>
								<tr>
									<th><label for="Width">사이즈조정</label></th>
									<td>
										가로 - <input type="text" id="Width" name="Width" title="가로사이즈" value="<%=Width%>" maxlength="10" class="w65" /> px /
										세로 - <input type="text" id="Height" name="Height" title="세로사이즈" value="<%=height%>" maxlength="10" class="w65" /> px
										 / ex)(318*300)
									</td>
								</tr>
								<tr>
									<th><label for="BDate">기간설정</label></th>
									<td>
										<input type="text" id="BDate" name="BDate" title="게시시작일" value="<%=BDate%>" readonly style="margin-right:5px;"/> 부터 
										~ <input type="text" id="EDate" name="EDate" title="게시종료일" value="<%=EDate%>" readonly style="margin-right:5px;" /> 까지
									</td>
								</tr>
								<tr>
									<th><label for="Contents">내용</label></th>
									<td>
										<textarea name="Contents" id="Contents" class="w660" style="display:none;"><%=Contents%></textarea>
										<textarea name="Contents_org" id="Contents_org" class="w660" style="display:none;"><%=Contents%></textarea>
									</td>
								</tr>

							</tbody>
						</table>
						</form>
					</div>
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_b"><a href="pop_list.asp?Page=<%=Page%>">목록보기</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
							</span><span class="btns btn_b"><a href="#btnPreview" id="btnPreview">미리보기</a>
							</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>