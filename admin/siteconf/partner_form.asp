<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : partner_form.asp / 파트너(입력업체, 온라인광고) 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Cate
	Dim PartnerName, BizNum, ChargeName, ChargeTel
	Dim strTitle, strTitleName
	
	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)

	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Cate = FN_Req("Cate","6")
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
		End If

		Select Case Cate
			Case "5" : strTitle = "온라인 광고"	: strTitleName = "온라인사이트명"
			Case "6" : strTitle = "인력업체"	: strTitleName = "인력업체명"
		End Select
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspPartner_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PartnerName = FN_InputVal(Rs("PartnerName"))
				BizNum 		= FN_InputVal(Rs("BizNum"))
				ChargeName 	= FN_InputVal(Rs("ChargeName"))
				ChargeTel 	= FN_InputVal(Rs("ChargeTel"))
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title><%=strTitle%> 등록 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			if (!$("#PartnerName").val()){
				alert("<%=strTitleName%>을 입력해주세요.");
				$("#PartnerName").focus();
				return;
			}
		
			$("#frmInfo").attr({action:"https://www.mpcjob.co.kr/admin/siteconf/partner_proc.asp", method:"post"}).submit();
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
			<h2><%=strTitle%> 등록</h2>
			<div class="tblArea">
				<form id="frmInfo" name="frmInfo" action="https://www.mpcjob.co.kr/admin/siteconf/partner_proc.asp" method="post">
				<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
				<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
				<input type="hidden" id="Cate" name="Cate" value="<%=Cate%>">
				<table class="tblForm">
					<colgroup>
						<col class="w95" />
						<col />
						<col class="w95" />
						<col />
					</colgroup>
					<tbody>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<tr>
							<th><label for="PartnerName"><%=strTitleName%></label></th>
							<td>
								<input type="text" id="PartnerName" name="PartnerName" title="<%=strTitleName%>" value="<%=PartnerName%>" maxlength="30" class="w330" /> 
							</td>
							<th><label for="BizNum">사업자등록번호</label></th>
							<td>
								<input type="text" id="BizNum" name="BizNum" title="사업자등록번호" value="<%=BizNum%>" maxlength="15" class="w330" style="ime-mode:disabled;" /> 
							</td>
						</tr>
						<tr>
							<th><label for="ChargeName">담당자 이름</label></th>
							<td>
								<input type="text" id="ChargeName" name="ChargeName" title="담당자 이름" value="<%=ChargeName%>" maxlength="20" class="w330" /> 
							</td>
							<th><label for="ChargeTel">담당자 전화번호</label></th>
							<td>
								<input type="text" id="ChargeTel" name="ChargeTel" title="담당자 전화번호" value="<%If ChargeTel <> "" Then Response.write oSeed.Decrypt(ChargeTel) End If %>" maxlength="20" class="w330" style="ime-mode:disabled;" /> 
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