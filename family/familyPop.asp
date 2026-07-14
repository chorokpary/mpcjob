<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : familyPop.asp / MPC Family 목록 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim Part
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strTitle

	Call Main()
	
	Sub Main()

		listPage 	= 5
		listRecord 	= 8	'목록갯수
		Page 		= FN_Req("Page","1")

		Part 	= FN_Req("Part","통신분야")
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Select Case Part
			Case "통신분야" : strTitle = "tit_familyPop_01.png"
			Case "공공기관" : strTitle = "tit_familyPop_02.png"
			Case "쇼핑분야" : strTitle = "tit_familyPop_03.png"
			Case "교육분야" : strTitle = "tit_familyPop_04.png"
			Case "금융분야" : strTitle = "tit_familyPop_05.png"
			Case "기타분야" : strTitle = "tit_familyPop_06.png"
			Case "IT분야" : strTitle = "tit_familyPop_07.png"
			Case "건설분야" : strTitle = "tit_familyPop_08.png"
		End Select
		
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspProject_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT_FAM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= "PART"
			.Parameters("@sWord").Value 	= Part
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), PrjSeq(1), PrjName(2), LogoPath(3), Flag(4)
			kFields 	= Array("TotalCnt","PrjSeq","PrjName","LogoPath","Flag")
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
	<div class="inner"><div class="inner2">
		<div class="popCont familypop">
			<h3><img src="/images/pages/<%=strTitle%>" alt="HC Family <%=Part%>" /></h3>
			<div class="familyList">
				<ul>
					<% If Not kData Then %>
					<li class="listnone">준비중입니다.</li>
					<% Else %>
					<% 		
							' TotalCnt(0), PrjSeq(1), PrjName(2), LogoPath(3), Flag(4)
							vNum = RecordCount - ((Page-1) * listRecord) 
							For i = 0 To arrDataNum 
					%>
					<li>
						<div class="imgs">
							<img src="<%=Fn_SetDefault(arrData(3,i),"","/images/common/no_company.jpg",arrData(3,i))%>" alt="" />
							<div class="bg"></div>
						</div>
						<p class="info"><%=arrData(2,i)%></p>
					</li>

					<%			vNum = vNum - 1
							Next 
					%>
					<% End If %>
				</ul>
				<div class="paging">
					<% Call SB_PagingWrite_LayerPop(listPage, RecordCount, PageCount) %>
				</div>
				<img src="/images/pages/btn_popClose.jpg" alt="닫기" class="family_close" id="btnClose" />
				<form name="frmPageAjax" id="frmPageAjax" method="post">
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="Part" name="Part" value="<%=Part%>">
				</form>
			</div>
		</div>
	</div></div>