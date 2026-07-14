<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : partner_list.asp / 파트너 목록 (인력업체 & 온라인광고관리)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strBtnAdd, strTitleName

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)


	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 40	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		Select Case CONF_CATE
			Case "5" : strBtnAdd = "온라인광고 등록"	: strTitleName = "온라인사이트"
			Case "6" : strBtnAdd = "인력업체 등록"		: strTitleName = "인력업체명"
		End Select
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspPartner_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sCate").Value 	= CONF_CATE
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@sView").Value 	= CONF_VIEW
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), Seq(1), PartnerName(2), BizNum(3), ChargeName(4), ChargeTel(5), IsView(6)
			kFields 	= Array("TotalCnt","Seq","PartnerName","BizNum","ChargeName","ChargeTel","IsView")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
			
			RecordCount = arrData(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
					<script type="text/javascript">
						function chkDel(){
							if (cntChecked("#frmList")==0){
								alert("선택된 데이터가 없습니다.");
							}else{
								if(confirm("선택한 데이터를 <%=Fn_SetDefault(CONF_VIEW,"1","삭제","복구")%>하시겠습니까?")){
									$("#Flag").val("STA");
									$("#frmList").attr({action:"partner_proc.asp", method:"post"}).submit();
								}
							}
						}
						
						function openForm(seq){
							openWin("partner_form.asp?Cate=<%=CONF_CATE%>&Seq="+seq,"popPartner",0,0,930,400,0)
						}
						
						
						$(function() {
						    $("#sLevel").change(initPage);
						    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
						    $("#btnStatus").click(chkDel);
						    $("#btnReg").click(function(){ openForm(0);});
						});
					</script>

					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnStatus" id="btnStatus"><%=Fn_SetDefault(CONF_VIEW,"1","삭제","복구")%></a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg"><%=strBtnAdd%></a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<input type="hidden" id="sKey" name="sKey" value="NAME">
							<input type="hidden" id="sView" name="sView" value="<%=CONF_VIEW%>">
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>" class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
						</form>
						</div>
					</div>
					
					<div class="tblArea">
						<form name="frmList" id="frmList">
						<input type="hidden" id="Flag" name="Flag">
						<input type="hidden" id="Cate" name="Cate" value="<%=CONF_CATE%>">
						<input type="hidden" id="sView" name="sView" value="<%=CONF_VIEW%>">
						<table <%=Fn_SetDefault(CONF_VIEW,"0","class=""closed""","")%>>
							<colgroup>
								<col class="w40" />
								<col class="w40" />
								<col />
								<col class="w125" span="3" />
							</colgroup>
							<thead>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th><%=strTitleName%></th>
									<th>사업자등록번호</th>
									<th>담당자 이름</th>
									<th>담당자 전화번호</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="6">등록된 데이터가 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), Seq(1), PartnerName(2), BizNum(3), ChargeName(4), ChargeTel(5), IsView(6)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td><a href="javascript:openForm('<%=arrData(1,i)%>')"><%=FN_ClearTag(arrData(2,i))%></a></td>
									<td><a href="javascript:openForm('<%=arrData(1,i)%>')"><%=arrData(3,i)%></a></td>
									<td><%=arrData(4,i)%></td>
									<td><%If arrData(5,i) <> "" Then Response.write oSeed.Decrypt(arrData(5,i)) End If %></td>
								</tr>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
						</form>
						<div class="paging">
							<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
						</div>
					</div>
