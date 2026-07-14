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
			.Parameters("@FLAG").Value 	= "MAIN2"
			
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
<!-- imgUrl = 이미지 경로 수정,  link = 링크 수정,  target = 링크 타겟 설정 -->
<% If arrDataNum > -1 Then %>
	<% For i = 0 To arrDataNum %>
	<A HREF="<%=arrData(2,i)%>" target="<%=Fn_SetDefault(InStr(arrData(2,i),"http://"),"0","_self","_blank")%>"><img SRC="<%=arrData(1,i)%>"></A>
	<% Next %>
<% End If %>