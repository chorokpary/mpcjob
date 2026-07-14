<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : hiretype_form.asp / 고용형태 등록 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2013-03-13 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Sort, Flag, CONF_TBL, CONF_FLD
	Dim CdKey, IsUse
	
	Call Main()
	
	Sub Main()

		CONF_TBL = "TBL_RECRUIT"
		CONF_FLD = "HireType"

		Sort = FN_Req("Sort","0")
		
		If Sort <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
			IsUse = True
		End If

	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Tbl").Value 	= CONF_TBL
			.Parameters("@Field").Value = CONF_FLD
			.Parameters("@Sort").Value 	= Sort
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				CdKey = FN_InputVal(Rs("CdKey"))
				IsUse = Rs("IsUse")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
						<script type="text/javascript">
							function chkForm(){
								if (!$("#Key").val()){
									alert("고용형태를 입력해주세요.");
									$("#Key").focus();
									return;
								}
							
								$("#frmInfo").attr({action:"hiretype_proc.asp", method:"post"}).submit();
							}
							
							$(function() {
								$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); $("#divInfo").hide(); })
								$("#btnSubmit").click(chkForm);
							});
						</script>

						<h2 style="margin-top:20px;">고용형태&nbsp;<%=Fn_SetDefault(Flag,"ADD","등록","수정")%></h2>
						
						<div class="tblArea">
							<form id="frmInfo" name="frmInfo" action="hiretype_proc.asp" method="post">
							<input type="hidden" id="Sort" name="Sort" value="<%=Sort%>">
							<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">

							<table class="tblForm">
								<colgroup>
									<col class="w100" />
									<col />
									<col class="w100" />
									<col class="w350" />
								</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr>
										<th><label for="Key">고용형태</label></th>
										<td>
											<input type="text" id="Key" name="Key" title="고용형태" value="<%=CdKey%>" class="w250" maxlength="25" /> 
										</td>
										<th><label for="IsUse_Y">사용여부</label></th>
										<td>
											<input type="radio" id="IsUse_Y" name="IsUse" value="1" title="선택된 고용형태 사용"<%=Fn_SetDefault(IsUse,"True"," checked","")%>> 
											<label for="IsUse_Y" title="선택된 고용형태 사용">사용</label>
											<input type="radio" id="IsUse_N" name="IsUse" value="0" title="선택된 고용형태 사용하지 않음"<%=Fn_SetDefault(IsUse,"False"," checked","")%>> 
											<label for="IsUse_N" title="선택된 고용형태 사용하지 않음">사용안함</label>
										</td>
									</tr>
								</tbody>
							</table>
							</form>
						</div>
						<% If Flag = "MOD" Then %>
						<p class="note_c">* 고용형태를 수정할 경우 등록 되어있는 채용 공고 정보도 함께 변경됩니다.</p>
						<% End If %>

						<div class="btnArea">
							<div class="fl_r">
								<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
								</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
							</div>
						</div>
