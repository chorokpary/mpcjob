<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : info.xml.asp / 플래시 Config
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim IsData
	
	Call Main()
	
	Sub Main()

		Call getData
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfFlash_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "MAIN"
			
			' ## 회원정보
			Set Rs = .Execute

			If Rs.Eof or Rs.Bof Then
				arrDataNum 	= -1
				kData 		= False
			Else
				' Seq(0), ImgPath(1), LinkUrl(2), Sort(3)
				kFields 	= Array("Seq","ImgPath","LinkUrl","Sort")
				arrData 	= Rs.GetRows(,,kFields)
				arrDataNum 	= UBound(arrData, 2)
				kData 		= True
			End If
			
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<?xml version="1.0" encoding="UTF-8"?>
<!-- xml경로는 Main.as파일 여시고 "info.xml"를 수정하시면 됩니다. -->
<ele>
<!-- imgUrl = 이미지 경로 수정,  link = 링크 수정,  target = 링크 타겟 설정 -->
<% If arrDataNum > -1 Then %>
	<% For i = 0 To arrDataNum %>
	<img url="<%=arrData(1,i)%>" link="<%=arrData(2,i)%>" target="<%=Fn_SetDefault(InStr(arrData(2,i),"http://"),"0","_self","_blank")%>">
		<contents menu="Event<%=i+1%>" title="<%=arrData(2,i)%>"/>
	</img>
	<% Next %>
<% End If %>

	<!--
		fontUrl            = 폰트 경로 설정(상대 경로로 설정해야 합니다.)
		fontSpace          = 텍스트 사이 간격

		imgZ               = 이미지 사이 간격
		imgGlowColor       = 이미지 테두리 색상
		imgGlowAlp         = 이미지 테두리 투명도 설정
		imgGlowScale       = 이미지 테두리 두께

		textBoxX           = 메뉴 텍스트 x좌표
		textBoxY           = 메뉴 텍스트 y좌표
		textX              = 메뉴 텍스트 menu라인 사이 간격
		textColor          = 메뉴 텍스트 색상
		textSize           = 메뉴 텍스트 폰트 사이즈

		titleBoxX          = 내용 텍스트 x좌표
		titleBoxY          = 내용 텍스트 y좌표
		titleBoxMenuSize   = 내용 텍스트 메뉴 폰트 사이즈
		titleSize          = 내용 텍스트 내용 폰트 사이즈

		slideTimer         = 자동 슬라이드 시간(초당 설정)
	-->
	<option
		fontUrl            = "/images/flash/font/font.swf"
		fontSpace          = "0"
		
		imgGlowColor       = "ffffff"
		imgGlowAlp         = "1"
		imgGlowScale       = "6"

		imgRotation        = "70"
		imgSpace           = "55"
		imgCenterSpace     = "100"
		imgRotate          = "4"
		
		leftBtnX           = "345"
		rightBtnX          = "655"
		btnY               = "435"

		menuX              = "451"
		menuY              = "441"
		menuColor          = "00D2FF"

		subX               = "451"
		subY               = "472"
		subColor           = "ffffff"

		Shadow             = "20"
		ShadowAlp          = "55"
		ShadowY            = "2"
		
		speed              = "0.5"
		checkDirect        = "left"
		slideTime          = "3"
	/>
</ele>