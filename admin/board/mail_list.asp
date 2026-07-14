<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mail_list.asp / SMS 발송 내역
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-30 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strTabFlag

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 15	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
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
			.CommandText ="uspSmsHistory_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), Seq(1), Subject(2), Contents(3), RegDate(4), PrjName(5), AdmName(6)

			kFields 	= Array("TotalCnt","Seq","Subject","Contents","RegDate","PrjName","AdmName")
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
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>메일링 서비스 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>메일링 서비스</h2>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="mail_form.asp">SMS 발송</a></span>
						</div>
						
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Seq" name="Seq">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<div class="fl_r">
								<select id="sKey" name="sKey" title="검색항목">
									<option value="">검색선택</option>
									<option value="CONTENT"<%=Fn_SetDefault(sKey,"CONTENT"," selected","")%>>내용</option>
									<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>보낸이</option>
								</select>
								<input type="text" id="sWord" name="sWord" value="<%=sWord%>" class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
							</div>
						</form>
					</div>
					
					<div class="tblArea">
						<table>
							<colgroup>
								<col class="w40" />
								<col class="w40" />
								<col />
								<col class="w110" />
								<col class="w75" />
								<col class="w135" />
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
									<th>내용</th>
									<th>프로젝트명</th>
									<th>담당자</th>
									<th>발송일</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="6">등록된 SMS 발송 내역이 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										' TotalCnt(0), Seq(1), Subject(2), Contents(3), RegDate(4), PrjName(5), AdmName(6)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<tr>
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td class="alignL"><%=FN_ClearTag(arrData(3,i))%></td>
									<td><%=arrData(5,i)%></td>
									<td><%=arrData(6,i)%></td>
									<td><%=arrData(4,i)%></td>
								</tr>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
						<div class="paging">
							<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>