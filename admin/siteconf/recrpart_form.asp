<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : recrpart_form.asp / 직종코드 등록 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2013-03-14 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Sort, Sort2, Flag
	Dim CdKey, IsUse, CdKey2, IsUse2
	Dim CONF_TBL, CONF_FLD, CONF_TBL2, CONF_FLD2
	
	Call Main()
	
	Sub Main()

		CONF_TBL = "TBL_RECRUIT"
		CONF_FLD = "RecrPart"

		CONF_TBL2 = "TBL_RECRUIT"
		CONF_FLD2 = "RecrPart2"

		Sort = FN_Req("Sort","0")
		Sort2 = FN_Req("Sort2","2")
		
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
			
			.Parameters("@Tbl2").Value 	= CONF_TBL2
			.Parameters("@Field2").Value = CONF_FLD2
			.Parameters("@Sort2").Value = Sort2
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				CdKey = FN_InputVal(Rs("CdKey"))
				IsUse = Rs("IsUse")
			End If
			
			If Sort2 <> "0" Then
				Set Rs = Rs.NextRecordSet()
				
				If Not Rs.Eof Then
					CdKey2 = FN_InputVal(Rs("CdKey"))
					IsUse2 = Rs("IsUse")
				End If
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
						<script type="text/javascript">
							function chkForm(){
								if (!$("#Key").val()){
									alert("1차 직종을 입력해주세요.");
									$("#Key").focus();
									return;
								}
								if (!$("#Key2").val()){
									alert("2차 직종을 입력해주세요.");
									$("#Key2").focus();
									return;
								}
							
								$("#frmInfo").attr({action:"recrpart_proc.asp", method:"post"}).submit();
							}
							
							function setKey(){

								if ($(this).val()){
									$("#Key").val($(this).val());
									$("#Key").addClass("lock");
									$("#Key").attr("readonly",true);
								}else{
									$("#Key").removeClass("lock");
									$("#Key").attr("readonly",false);
								}
							}
							
							$(function() {
								$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); $("#divInfo").hide(); });
								$("#btnSubmit").click(chkForm);
								$("#selKey").change(setKey);
							});
						</script>

						<h2 style="margin-top:20px;">직종&nbsp;<%=Fn_SetDefault(Flag,"ADD","등록","수정")%></h2>
						
						<div class="tblArea">
							<form id="frmInfo" name="frmInfo" action="recrpart_proc.asp" method="post">
							<input type="hidden" id="Sort" name="Sort" value="<%=Sort%>">
							<input type="hidden" id="Sort2" name="Sort2" value="<%=Sort2%>">
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
										<th <%=Fn_SetDefault(Flag,"MOD"," style=""height:50px;""","")%>><label for="Key">1차 직종</label></th>
										<td>
											<% Call SB_MakeHTMLSelect_CdTbl_Full(objDBCon, CONF_TBL, CONF_FLD, "selKey", ":: 직접입력 ::", "", CdKey, "", " class=""w150"" ", False, "", "ALL")%>
											<input type="text" id="Key" name="Key" title="1차 직종" value="<%=CdKey%>" class="w150<%=Fn_SetDefault(Flag,"MOD"," lock","")%>" maxlength="25"<%=Fn_SetDefault(Flag,"MOD"," readonly","")%> /> 
											<% If Flag = "MOD" Then %>
											<p style="margin-top :5px;">
												<input type="checkbox" name="ChangeAll" id="ChangeAll" Value="T" title="동일한 1차 직종 전체  변경">
												<label for="ChangeAll">동일한 1차 직종 전체  변경</label>
											</p>
											<% End If %>
										</td>
										<th><label for="Key2">2차 직종</label></th>
										<td>
											<input type="text" id="Key2" name="Key2" title="2차 직종" value="<%=CdKey2%>" class="w250" maxlength="25" /> 
										</td>
									</tr>
									<tr>
										<th><label for="IsUse_Y">사용여부</label></th>
										<td colspan="3">
											<input type="radio" id="IsUse_Y" name="IsUse" value="1" title="선택된 직종 사용"<%=Fn_SetDefault(IsUse,"True"," checked","")%>> 
											<label for="IsUse_Y" title="선택된 직종 사용">사용</label>
											<input type="radio" id="IsUse_N" name="IsUse" value="0" title="선택된 직종 사용하지 않음"<%=Fn_SetDefault(IsUse,"False"," checked","")%>> 
											<label for="IsUse_N" title="선택된 직종 사용하지 않음">사용안함</label>
										</td>
									</tr>
								</tbody>
							</table>
							</form>
						</div>
						<% If Flag = "MOD" Then %>
						<p class="note_c">* 직종을 수정할 경우 등록 되어있는 채용 공고 정보도 함께 변경됩니다.</p>
						<% End If %>

						<div class="btnArea">
							<div class="fl_r">
								<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
								</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
							</div>
						</div>
