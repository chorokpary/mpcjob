<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : faq_form.asp / FAQ 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-02 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim Subject, Attach1_Path, Attach1_Name, Attach2_Path, Attach2_Name, Contents, Category
	
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
			.Parameters("@UseCnt").Value = 0
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				Subject = FN_InputVal(Rs("Subject"))
				Attach1_Path = Rs("Attach1_Path")
				Attach1_Name = Rs("Attach1_Name")
				Attach2_Path = Rs("Attach2_Path")
				Attach2_Name = Rs("Attach2_Name")
				Category = Rs("Category")
				Contents = FN_BrToNL(Rs("Contents"))
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
	<title>FAQ 등록 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			if (!$("#Subject").val()){
				alert("질문을 입력해주세요.");
				$("#Subject").focus();
				return;
			}
			if (!$("#Category").val()){
				alert("분류를 선택해주세요.");
				$("#Category").focus();
				return;
			}
			if (!$("#Contents").val()){
				alert("답변을 입력해주세요.");
				$("#Contents").focus();
				return;
			}
			
			$("#frmInfo").attr({action:"faq_proc.asp", method:"post"}).submit();
		}
		
		$(function() {
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); $("#pAttach1").show(); $("#pAttach2").show(); })
			$("#btnSubmit").click(chkForm);
		});
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<h2>FAQ 등록</h2>
			<div class="tblArea">
			
				<form id="frmInfo" name="frmInfo" action="faq_proc.asp" method="post" enctype="multipart/form-data" >
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
							<th><label for="">질문</label></th>
							<td>
								<input type="text" id="Subject" name="Subject" Title="질문" value="<%=Subject%>" maxlength="100" class="w780" /> 
							</td>
						</tr>
						<tr>
							<th><label for="">분류</label></th>
							<td>
								<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_BBS", "Category_faq", "Category", "분류", "", Category, "등록된 카테고리가 없습니다.", " class=""w290"" ", False)%>
							</td>
						</tr>
						<tr>
							<th><label for="">답변</label></th>
							<td>
								<textarea id="Contents" name="Contents" ><%=Contents%></textarea>
							</td>
						</tr>
					</tbody>
				</table>
				
				</form>
			</div>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
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