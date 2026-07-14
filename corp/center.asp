<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : center.asp / 센터찾기
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Center, PrjSeq
	Dim kData, kFields, i, vNum
	Dim arrTab, arrTabNum, arrPrj, arrPrjNum
	
	Dim IsData
	Dim PrjName, LogoPath, MapPath, Addr, Part, Flag, TransSubway, TransBus, Location

	Call Main()
	
	Sub Main()

		Center 	= FN_Req("Center","본사")
		PrjSeq 	= FN_Req("PrjSeq","0")
		
		Call getTabMenu()
		Call getPrjList()
		Call getPrjInfo()
		
	End Sub

	Sub getTabMenu()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "SEL"
			.Parameters("@Tbl").Value 	= "TBL_PROJECT"
			.Parameters("@Field").Value = "Center"
			
			' ## 코드 정보
			Set Rs = .Execute


		End With
			
		If Rs.eof Then
			arrTabNum = -1
		Else
			arrTab = Rs.getrows
			arrTabNum = ubound(arrTab,2)
		End if
		
		Rs.Close
		Set Rs = Nothing

	End Sub
	
	Sub getPrjList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspProject_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "CTLST"
			.Parameters("@Center").Value 	= Center
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrPrjNum 	= -1
		Else
			' PrjSeq(0), PrjName(1)
			kFields 	= Array("PrjSeq","PrjName")
			arrPrj 	= Rs.GetRows(,,kFields)
			arrPrjNum 	= UBound(arrPrj, 2)
			
			If PrjSeq = "0" Then
				PrjSeq = arrPrj(0,0)
			End If
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

	Sub getPrjInfo()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspProject_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "INFO"
			.Parameters("@Seq").Value 	= PrjSeq
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			IsData =False
		Else
			IsData =True
			
			PrjName 	= Rs("PrjName")
			LogoPath 	= Rs("LogoPath")
			MapPath 	= Rs("MapPath")
			Addr 		= Rs("Addr")
			Part 		= Rs("Part")
			Flag 		= Rs("Flag")
			TransSubway = Rs("TransSubway")
			TransBus 	= Rs("TransBus")
			Location 	= Rs("Location")
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>About HC - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
</head>
<body>
	<div id="wrap" class="office">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_corp.jpg" alt="회사소개" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<% If arrTabNum > -1 Then %>
				<div class="tabArea">
					<% For i = 0 To arrTabNum%><span class="tab<%=Fn_SetDefault(Center,arrTab(0,i)," on","")%>"><a href="center.asp?Center=<%=arrTab(0,i)%>"><%=arrTab(1,i)%></a></span><% Next %>
				</div>
				<% End If %>
				<div class="section">
					<div class="officeT">
						<div class="offceInnerT">
							<ul>
								<li>
									<% If arrPrjNum = -1 Then %>
									서비스센터가 없습니다.
									<% Else %>
										<% For i = 0 To arrPrjNum%><a href="center.asp?Center=<%=Center%>&PrjSeq=<%=arrPrj(0,i)%>" <%=Fn_SetDefault(PrjSeq,arrPrj(0,i)," class=""on""","")%>><%=arrPrj(1,i)%></a><% Next %>
									<% End If %>
								</li>
							</ul>
						</div>
					</div>
					<div class="officeB">
						<div class="offceInnerB">
						<% If IsData Then %>
							<% If MapPath <> "" Then %>
							<img src="<%=MapPath%>" alt="" width!="726" />
							<% End If %>
							<ul>
								<% If Addr <> "" Then %>
								<li class="home">상세주소
									<p><%=Addr%></p>
								</li>
								<% End If %>
								<% If TransBus <> "" Or TransSubway <> "" Then %>
								<li class="traffic">교통편
									<% If TransBus <> "" Then %><p>버스 : <%=TransBus%></p><% End If %>
									<% If TransSubway <> "" Then %><p>지하철 : <%=TransSubway%></p><% End If %>
								</li>
								<% End If %>
								<% If Location <> "" Then %>
								<li class="location">오시는길
									<p><%=FN_NLToBr(Location)%></p>
								</li>
								<% End If %>
							</ul>
						<% Else %>
							<ul>
								<li class="imgReady"><img src="/images/pages/map_ready.jpg" alt="이미지준비중입니다." /></li>
							</ul>
						<% End If %>
						</div>
					</div>
				</div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>