<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : adm_form.asp / 관리자 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag
	Dim AdmID, AdmName, AdmLevel, AdmTel, AdmMobile, AdmEmail
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		
		If Session("ALevel") = "3" And CStr(Session("ASeq")) <> CStr(Seq) Then
			Call SB_ReturnErr("[권한] 잘못된 경로로 접근하셨습니다.","CLOSE")
		End If
		
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
			.CommandText ="uspAdm_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@AdmSeq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				AdmID 		= FN_InputVal(Rs("AdmID"))
				AdmName 	= FN_InputVal(Rs("AdmName"))
				AdmLevel 	= Rs("AdmLevel")
				AdmTel 		= FN_InputVal(Rs("AdmTel"))
				AdmMobile 	= FN_InputVal(Rs("AdmMobile"))
				AdmEmail 	= FN_InputVal(Rs("AdmEmail"))
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>관리자 추가 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			<% If Session("ALevel") = "1" Then %>
			if (!$("#AdmLevel").val()){
				alert("관리자 구분을 선택해주세요.");
				$("#AdmLevel").focus();
				return;
			}
			<% End If %>
			if (!$("#AdmID").val()){
				alert("아이디를 입력해주세요.");
				$("#AdmID").focus();
				return;
			}
			<% If Flag = "ADD" Then %>
			if (!$("#AdmPwd").val()){
				alert("비밀번호를 입력해주세요.");
				$("#AdmPwd").focus();
				return;
			}
			<% End If %>
			if (!$("#AdmName").val()){
				alert("이름을 입력해주세요.");
				$("#AdmName").focus();
				return;
			}
		
			$("#frmInfo").attr({action:"adm_proc.asp", method:"post"}).submit();
		}
		
		$(function() {
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); })
			$("#btnSubmit").click(chkForm);
		});
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<h2>관리자 등록 (관리자를 추가하거나 수정하실 수 있습니다.)</h2>
			<div class="tblArea">
				<form id="frmInfo" name="frmInfo" action="adm_action.asp" method="post">
				<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
				<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
				<table class="tblForm">
					<colgroup>
						<col class="w85" />
						<col />
						<col class="w85" />
						<col />
					</colgroup>
					<tbody>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<% If Session("ALevel") = "1" Then %>
						<tr>
							<th><label for="">관리자 구분</label></th>
							<td colspan="3">
								<select id="AdmLevel" name="AdmLevel" title="관리자 구분" class="w340">
									<option value="">선택</option>
									<option value="1"<%=Fn_SetDefault(AdmLevel,"1"," selected","")%>>최고관리자</option>
									<option value="2"<%=Fn_SetDefault(AdmLevel,"2"," selected","")%>>최고관리자2</option>
									<option value="3"<%=Fn_SetDefault(AdmLevel,"3"," selected","")%>>일반관리자</option>
								</select> 
							</td>
						</tr>
						<% End If %>
						<tr>
							<th><label for="AdmID">아이디</label></th>
							<td>
								<input type="text" id="AdmID" name="AdmID" title="아이디" maxlength="30" value="<%=AdmID%>" class="w335" style="ime-mode:disabled;" <%=Fn_SetDefault(Flag,"MOD","readonly","")%> /> 
							</td>
							<th><label for="AdmPwd">비밀번호</label></th>
							<td>
								<input type="password" id="AdmPwd" name="AdmPwd" title="비밀번호" maxlength="30" value="" class="w335" /> 
							</td>
						</tr>
						<tr>
							<th><label for="AdmName">이름</label></th>
							<td>
								<input type="text" id="AdmName" name="AdmName" title="이름" maxlength="15" value="<%=AdmName%>" class="w335" /> 
							</td>
							<th><label for="AdmEmail">이메일</label></th>
							<td>
								<input type="text" id="AdmEmail" name="AdmEmail" title="이메일" maxlength="50" value="<%=AdmEmail%>" class="w335" style="ime-mode:disabled;" /> 
							</td>
						</tr>
						<tr>
							<th><label for="AdmTel">연락처</label></th>
							<td>
								<input type="text" id="AdmTel" name="AdmTel" title="연락처" maxlength="15" value="<%=AdmTel%>" class="w335" style="ime-mode:disabled;" /> 
							</td>
							<th><label for="AdmMobile">휴대폰</label></th>
							<td>
								<input type="text" id="AdmMobile" name="AdmMobile" title="휴대폰" maxlength="15" value="<%=AdmMobile%>" class="w335" style="ime-mode:disabled;" /> 
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