<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : conf_main.asp / 메인설정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-22 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim ConfGreeting, ConfFormDoc, ConfFormXls, ConfFormHwp
	
	Call Main()
	
	Sub Main()

		Call getData
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfSite_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				ConfGreeting = FN_InputVal(Rs("ConfGreeting"))
				ConfFormDoc = Rs("ConfFormDoc")
				ConfFormXls = Rs("ConfFormXls")
				ConfFormHwp = Rs("ConfFormHwp")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>메인설정 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function doSubmit(){
			$("#frmInfo").attr({action:"conf_main_proc.asp", method:"post"}).submit();
		}
		
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}
		
		$(function() {
			$("#btnSubmit1").click(doSubmit);
			$("#btnSubmit2").click(doSubmit);
			$("#btnConfFlash").click(function(){ openWin("conf_flash.asp","popConfFlash",0,0,930,630,0); });
			$("#btnRecom").click(function(){ openWin("../recruit/job_recom_list.asp?sViewFlag=MAIN","popRecom",0,0,930,630,0); });
		});
	</script>
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				
				<form id="frmInfo" name="frmInfo" action="pjt_proc.asp" method="post" enctype="multipart/form-data" >
				<div id="contArea">
					<h2>인사말 관리</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="ConfGreeting">HEAD 인사말관리</label></th>
									<td>
										<input type="text" id="ConfGreeting" name="ConfGreeting" title="인사말 관리" value="<%=ConfGreeting%>" class="w550" maxlength="200" /> <span class="btns btn_b"><a href="#" id="btnSubmit1">수정하기</a></span>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>FLASH 관리</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="">FLASH 관리</label></th>
									<td>
										<span class="btns btn_b"><a href="#btnConfFlash" id="btnConfFlash">수정하기</a></span>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>이달의 추천공고 관리</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="">이달의 추천공고</label></th>
									<td>
										<span class="btns btn_b"><a href="#btnRecom" id="btnRecom">수정하기</a></span>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>이력서 양식 다운로드 양식 등록</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="ConfFormDoc">WORD</label></th>
									<td>
										<% If Not FN_isBlank(ConfFormDoc) Then %>
										<p id="pConfFormDoc"><%=MID(ConfFormDoc, InStrRev(ConfFormDoc,"/") + 1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pConfFormDoc','DelConfFormDoc','<%=ConfFormDoc%>')">삭제</a></span></p>
										<% End If %>
										<input type="file" id="ConfFormDoc" name="ConfFormDoc" title="워드 이력서" value="" class="w660" size="95" />
										<input type="hidden" id="DelConfFormDoc" name="DelConfFormDoc" value="">
									</td>
								</tr>
								<tr>
									<th><label for="ConfFormXls">EXCEL</label></th>
									<td>
										<% If Not FN_isBlank(ConfFormXls) Then %>
										<p id="pConfFormXls"><%=MID(ConfFormXls, InStrRev(ConfFormXls,"/") + 1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pConfFormXls','DelConfFormXls','<%=ConfFormXls%>')">삭제</a></span></p>
										<% End If %>
										<input type="file" id="ConfFormXls" name="ConfFormXls" title="엑셀 이력서" value="" class="w660" size="95" />
										<input type="hidden" id="DelConfFormXls" name="DelConfFormXls" value="">
									</td>
								</tr>
								<tr>
									<th><label for="ConfFormHwp">HWP</label></th>
									<td>
										<% If Not FN_isBlank(ConfFormHwp) Then %>
										<p id="pConfFormHwp"><%=MID(ConfFormHwp, InStrRev(ConfFormHwp,"/") + 1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pConfFormHwp','DelConfFormHwp','<%=ConfFormHwp%>')">삭제</a></span></p>
										<% End If %>
										<input type="file" id="ConfFormHwp" name="ConfFormHwp" title="한글 이력서" value="" class="w660" size="95" />
										<input type="hidden" id="DelConfFormHwp" name="DelConfFormHwp" value="">
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					
					<div class="btnArea">
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnSubmit2">등록하기</a></span>
						</div>
					</div>
				</div>
				</form>
				
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>