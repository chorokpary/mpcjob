<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : notice_form.asp / 공지사항 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-29 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim Subject, Attach1_Path, Attach1_Name, Attach2_Path, Attach2_Name, Contents
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspBbs_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@BbsSeq").Value = Seq
			.Parameters("@BbsKey").Value = "notice"
			.Parameters("@UseCnt").Value = 0
			
			' ## 공지사항 정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				Subject = FN_InputVal(Rs("Subject"))
				Attach1_Path = Rs("Attach1_Path")
				Attach1_Name = Rs("Attach1_Name")
				Attach2_Path = Rs("Attach2_Path")
				Attach2_Name = Rs("Attach2_Name")
				Contents = Rs("Contents")
				'FN_setBlank
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>공지사항 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript" src="/common/Dev/SmartEditor/js/HuskyEZCreator.js" charset="utf-8"></script>
	<script type="text/javascript">
	<!--
		var oEditors = [];
		
		function chkForm(){
			if (!$("#Subject").val()){
				alert("제목을 입력해주세요.");
				$("#Subject").focus();
				return;
			}
			
			oEditors.getById["Contents"].exec("UPDATE_CONTENTS_FIELD", []);
		
			$("#frmInfo").attr({action:"notice_proc.asp", method:"post"}).submit();
		}
		
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}
		
		function doReset(){
			$("#frmInfo")[0].reset(); 
			$("#pAttach1").show(); 
			$("#pAttach2").show();
			
			var content = $("#Contents_org").val();
			oEditors.getById["Contents"].exec("SET_IR", [content]);
		}
		
		$(function() {
			// ### WYSIWYG
			nhn.husky.EZCreator.createInIFrame({
				oAppRef: oEditors,
				elPlaceHolder: "Contents",
				sSkinURI: "/common/Dev/SmartEditor/SmartEditor2Skin.html",	
				fCreator: "createSEditor2"
			});

			$("#btnCancel").click(doReset);
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
					<h2>공지사항</h2>

					<div class="tblArea">
					
						<form id="frmInfo" name="frmInfo" action="pjt_proc.asp" method="post" enctype="multipart/form-data" >
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
									<th><label for="Subject">제목</label></th>
									<td>
										<input type="text" id="Subject" name="Subject" title="제목" value="<%=Subject%>" maxlength="100" class="w660" /> 
									</td>
								</tr>	
								<tr>
									<th><label for="Contents">내용</label></th>
									<td>
										<textarea name="Contents" id="Contents" class="w660" style="display:none;"><%=Contents%></textarea>
										<textarea name="Contents_org" id="Contents_org" class="w660" style="display:none;"><%=Contents%></textarea>
									</td>
								</tr>	
								<tr>
									<th><label for="">첨부파일 1</label></th>
									<td>
										<% If Not FN_isBlank(Attach1_Name) Then %>
										<p id="pAttach1"><%=Attach1_Name%> <span class="btns btn_in"><a href="javascript:doDelFile('pAttach1','DelAttach1','<%=Attach1_Path%>')">삭제</a></span></p>
										<% End If %>
										<input type="file" id="Attach1" name="Attach1" title="첨부파일 1" value="" class="w660" size="95" /> 
										<input type="hidden" id="DelAttach1" name="DelAttach1" value="">
									</td>
								</tr>
								<tr>
									<th><label for="">첨부파일 2</label></th>
									<td>
										<% If Not FN_isBlank(Attach2_Name) Then %>
										<p id="pAttach2"><%=Attach2_Name%> <span class="btns btn_in"><a href="javascript:doDelFile('pAttach2','DelAttach2','<%=Attach2_Path%>')">삭제</a></span></p>
										<% End If %>
										<input type="file" id="Attach2" name="Attach2" title="첨부파일 2" value="" class="w660" size="95" /> 
										<input type="hidden" id="DelAttach2" name="DelAttach2" value="">
									</td>
								</tr>
							</tbody>
						</table>
						</form>
					</div>
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_b"><a href="notice_list.asp?Page=<%=Page%>">목록보기</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
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