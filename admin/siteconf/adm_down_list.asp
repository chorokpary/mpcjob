<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_apply_user.asp / 지원자 관리 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.Buffer = True '버퍼 사용

	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim sStatus, sBDate, sEDate
	Dim SortKey, SortDIr
	Dim AdmSeq

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)


	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= FN_Req("ListSize","100")	'목록갯수
		Page 		= FN_Req("Page","1")

		AdmSeq 		= FN_Req("AdmSeq","0")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))

		If UCase(sKey) = "MOBILE" Then 
			sWord = oSeed.Encrypt(sWord)
		End If 

		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 100
		End If
		
		Call getList()
		
	End Sub
	
	
	Sub getList()

		Dim Rs
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspAdm_Down_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@AdmSeq").Value 	= AdmSeq
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
			' TotalCnt(0), AdmLSeq(1), AdmSeq(2), RegIp(3), RegDate(4), Type(5), Etc1(6), Etc1(7)
			kFields 	= Array("TotalCnt","AdmLSeq","AdmSeq","RegIp","RegDate","Type","Etc1","Etc1")
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
	<title>다운로드 리스트 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body class="pop w1020">
	<div id="wrap" class="applicant recruitManage">
		<div id="contArea">
			<h2>다운로드 리스트 </h2>
			<% Response.Flush %>
			<div class="inner">
				<form name="frmList" id="frmList">
				<input type="hidden" id="Flag" name="Flag">
				<div class="tblArea">
					<table>
						<colgroup>
							<col class="w50" />
							<col class="w200" />
							<col class="w200" />
							<col />
						</colgroup>
						<thead>
							<tr class="design">
								<td></td>
								<td></td>
								<td></td>
								<td></td>
							</tr>
							<tr>
								<th>번호</th>
								<th>다운로드 종류</th>
								<th>접속 IP</th>
								<th>다운로드 날짜</th>
							</tr>
						</thead>
						<tbody>
							<% If Not kData Then %>
							<tr>
								<td colspan="4">다운로드 리스트가 없습니다.</td>
							</tr>
							<% Else %>
							<% 		

									' TotalCnt(0), AdmLSeq(1), AdmSeq(2), RegIp(3), RegDate(4), Type(5), Etc1(6), Etc1(7)						
									vNum = RecordCount - ((Page-1) * listRecord) 
									For i = 0 To arrDataNum 
										
							%>
							<tr>
								<td><%=vNum%></td>
								<td>
									<%
										Select Case arrData(5,i) 
											Case "R": Response.Write "채용정보" 
											Case "M" : Response.Write "회원정보" 
											Case Else : Response.Write "-" 
										End Select 
									%>
								</td>
								<td><%=arrData(3,i)%></td>
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
				</form>
			</div>
			<div class="btnArea bt1 pt10">
				<div class="fl_r">
					<span class="btns btn_g"><a href="javascript:;" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>

</body>
</html>