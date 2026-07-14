<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : matchArea.asp / 인근지역 채용공고 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sAddrSido, sAddrGugun
	Dim kData, kFields, i, vNum
	Dim arrList, arrListNum
	
	Dim strMsg
	
	Call Main()
	
	Sub Main()

		listPage 	= 5
		listRecord 	= 15	'목록갯수
		Page 		= FN_Req("Page","1")

		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getAreaInfo()
		
	End Sub

	Sub getAreaInfo()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "SIMP"
			.Parameters("@Seq").Value 		= Session("USeq")
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrListNum 	= -1
		Else
			sAddrSido = Rs("AddrSido") 
			sAddrGugun = Rs("AddrGugun")
			
			If sAddrSido = "" And sAddrGugun = "" Then
				strMsg = "이력서 관리에서 주소를 입력하셔야 인근지역공고 조회가 가능합니다."
			Else
				Call getList()
			End If
			
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub
		
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sAddrSido").Value = sAddrSido
			.Parameters("@sAddrGugun").Value = sAddrGugun

			.Parameters("@sViewFlag").Value = ""	' 전체글
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrListNum 	= -1
			kData 		= False
			strMsg = "해당지역에 일치하는 채용공고가 없습니다."
		Else
			' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","PrjName","LogoPath","AppEDate")
			arrList 	= Rs.GetRows(,,kFields)
			arrListNum 	= UBound(arrList, 2)
			kData 		= True

			RecordCount = arrList(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

	

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>MY PAGE - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function goView(seq){
			goViewCommon(seq,"/recruit/recruit_view.asp");
		}
	</script>
</head>
<body>
	<div id="wrap" class="matchArea">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_matchArea.jpg" alt="인근지역공고" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="normal">
					<ul class="pjtList">
						<% If Not kData Then %>
						<li class="nolist"><%=strMsg%></li>
						<% Else %>
						<% 		
								' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
								vNum = RecordCount - ((Page-1) * listRecord) 
								
								For i = 0 To arrListNum 
						%>
						<li>
							<a href="javascript:goView(<%=arrList(1,i)%>);">
								<div class="imgs">
									<img alt="<%=arrList(3,i)%>" src="<%=Fn_SetDefault(arrList(4,i),"","/images/common/no_company.jpg",arrList(4,i))%>">
									<div class="bg"></div>
								</div>
							</a>
							<div class="info">
								<h4><a href="javascript:goView(<%=arrList(1,i)%>);"><%=arrList(3,i)%></a></h4>
								<p class="pjt"><%=arrList(2,i)%></p>
								<p class="due">서류마감 : <%=FN_SetDateTimeFormat(arrList(5,i),"MM/DD")%></p>
							</div>
						</li>
						<%			vNum = vNum - 1
								Next 
						%>
						<% End If %>
					</ul>
					<div class="paging">
						<% Call SB_PagingWrite(listPage, RecordCount, PageCount) %>
					</div>

					<form name="frmPage" id="frmPage" method="post">
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="Seq" name="Seq">
					</form>

				</div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>