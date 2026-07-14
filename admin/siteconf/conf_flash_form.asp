<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_flash_form.asp / 플래시 관리 - 등록 / 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, sViewFlag
	Dim Seq, ImgPath, LinkUrl, Sort, BDate, EDate
	
	Call Main()
	
	Sub Main()
	
		Flag = FN_Req("Flag","ADD")
		sViewFlag = FN_Req("sViewFlag","ING")
		'Sort = FN_Req("Sort","1")
		Seq = FN_Req("Seq","0")

		If Seq <> "0" Then
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfFlash_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 이미지정보
			Set Rs = .Execute

			If Rs.Eof or Rs.Bof Then
				'Flag = "ADD"
			Else
				'Flag = "MOD"
				Sort = Rs("Sort")
				ImgPath = Rs("ImgPath")
				LinkUrl = FN_InputVal(Rs("LinkUrl"))
				BDate = Rs("BDate")
				EDate = Rs("EDate")
			End If
			
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>플래시 관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript">
		function doSubmit(){
			if (!$("#LinkUrl").val()){
				alert("LINK URL을 입력해주세요.");
				$("#LinkUrl").focus();
				return;
			}
			<% If Flag = "ADD" Then %>
			if (!$("#ImgPath").val()){
				alert("이미지를 등록해주세요.");
				$("#ImgPath").focus();
				return;
			}
			<% Else %>
			if ($("#DelImgPath").val() && !$("#ImgPath").val()){
				alert("이미지를 등록해주세요.");
				$("#ImgPath").focus();
				return;
			}
			<% End If %>
			$("#frmInfo").attr({action:"conf_flash_proc.asp", method:"post"}).submit();
		}
		
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}
		
		$(function() {
			$("#btnSubmit").click(doSubmit);
			$("#btnCancel").click(function(){ /*$("#frmInfo")[0].reset();*/ history.back(); });

			$("#BDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#EDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});

		 });
	</script>
</head>

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<h2>플래시 관리</h2>
			
			<form id="frmInfo" name="frmInfo" action="pjt_proc.asp" method="post" enctype="multipart/form-data" >
			<input type="hidden"  id="Flag" name="Flag" value="<%=Flag%>">
			<input type="hidden"  id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
			<input type="hidden"  id="Sort" name="Sort" value="<%=Sort%>">
			<input type="hidden"  id="Seq" name="Seq" value="<%=Seq%>">
			<!-- 재등록용 //-->
			<input type="hidden" id="TmpImg" name="TmpImg" value="<%=ImgPath%>">

			<div class="tblArea">
				<table class="tblForm">
					<colgroup>
						<col class="w100" />
						<col class="w120" />
						<col class="w100" />
						<col />
					</colgroup>
					<thead>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
					</thead>
					<tbody>
						<tr>
							<th>표시순서</th>
							<td><%=Sort%></td>
							<th>LINK URL</th>
							<td><input type="text" id="LinkUrl" name="LinkUrl" title="LinkUrl" value="<%=LinkUrl%>" class="w550" maxlength="100" /></td>
						</tr>
						<tr>
							<th>이미지 등록</th>
							<td colspan="3">
								<% If Not FN_isBlank(ImgPath) Then %>
								<p id="pImgPath"><%=MID(ImgPath, InStrRev(ImgPath,"/") + 1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pImgPath','DelImgPath','<%=ImgPath%>')">삭제</a></span></p>
								<% End If %>
								<input type="file" id="ImgPath" name="ImgPath" title="ImgPath" value="" class="w675" size="95" /> 
								<div class="mt5">(최적사이즈 : 339px * 295px / 업로드 가능 확장자 : jpg, gif, png, swf)</div>
								<input type="hidden" id="DelImgPath" name="DelImgPath" value="">
							</td>
						</tr>
						<tr>
							<th>게시기간</th>
							<td colspan="3">
								<input type="text" id="BDate" name="BDate" value="<%=BDate%>" style="margin-right:5px;" maxlength="10" /> 부터 ~ 
								<input type="text" id="EDate" name="EDate" value="<%=EDate%>" style="margin-right:5px;" maxlength="10" /> 까지
								(시작일과 종료일 모두 입력하셔야 게시기간이 반영됩니다.)
							</td>
						</tr>
					</tbody>
				</table>
			</div>
			</form>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#" id="btnSubmit">등록하기</a>
					</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">x 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>