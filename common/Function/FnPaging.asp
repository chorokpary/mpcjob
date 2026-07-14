<%
'____________________________________________________________________________________
'
' * Discription : FnPaging.asp / 게시판 페이징 함수
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2007-12-21 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


	' ########## 페이지 카운트
	Function FN_PageCount (argTotal, argPage)
		If argPage = 0 Then	
			FN_PageCount = 1
			'Response.End
		Else
			If argTotal mod argPage > 0 Then
				Dim tmpPgCnt : tmpPgCnt = FormatNumber(argTotal/ argPage,2)
				FN_PageCount =  CInt(Left(tmpPgCnt,InStr(tmpPgCnt,".")-1)) + 1
			Else
				FN_PageCount =  argTotal/ argPage
			End If
		End If
	End Function


	' ########## 페이징 출력 : Front 용
  	Sub	SB_PagingWrite (listPage, RecordCount, PageCount)	
		Dim	StartPage, EndPage,	FirstImg, PrevImg, NextImg,	LastImg
		Dim tmpCss : tmpCss = ""
		
		If PageCount <> "" And PageCount > 1 Then
		
			FirstImg = "<img src=""/images/common/btn_first.jpg"" alt=""첫페이지"">"
			PrevImg	= "<img src=""/images/common/btn_prev.jpg"" alt=""이전페이지"">"
			NextImg	= "<img src=""/images/common/btn_next.jpg"" alt=""다음페이지"">"
			LastImg	= "<img src=""/images/common/btn_last.jpg"" alt=""마지막페이지"">"
	
				
			StartPage =	Int	( (	Page - 1) /	listPage ) * listPage +	1

			EndPage	= (	StartPage +	listPage ) - 1
			If EndPage > PageCount Then
				EndPage	= PageCount	
			End	If

		
			If Page > 1	Then
				Response.Write "<a href=""JavaScript:goPage(1);"" class=""first"">" &	FirstImg & "</a> "
			Else
				Response.Write "<a class=""first"">" & FirstImg & "</a> "
			End	If

			If StartPage > 1 Then
				Response.Write "<a href=""JavaScript:goPage("	& StartPage - 1 & ");"" class=""prev"">" & PrevImg & "</a> "
			Else
				Response.Write "<a class=""prev"">" & PrevImg & "</a> "
			End	If
			
			For	i =	StartPage to EndPage
				tmpCss = ""
				
				if i = StartPage Then tmpCss = "num1"
				
				If i = Page	Then
					tmpCss = tmpCss & " on"
					Response.Write "<a class=""" & tmpCss &  """>"	& i	& "</a>	" 
				Else
					Response.Write "<a href=""JavaScript:goPage(" &	i &	");"" class=""" & tmpCss & """>" & i & "</a> "
				End	If 
			Next

			If EndPage < PageCount	Then
				Response.Write "<a href=""JavaScript:goPage("	& EndPage + 1 & ");"" class=""next"">" & NextImg & "</a> "
			Else
				Response.Write "<a class=""next"">" & NextImg & "</a> "
			End	If
			
			If PageCount > Page Then
				Response.Write "<a href=""JavaScript:goPage(" &	PageCount	& ");"" class=""last"">" &	LastImg	& "</a>"
			Else
				Response.Write "<a class=""last"">" & LastImg & "</a>"
			End If
		End If 
  	End	Sub

	' ########## 페이징 출력 : Front - Ajax용
  	Sub	SB_PagingWrite_LayerPop (listPage, RecordCount, PageCount)	
		Dim	StartPage, EndPage,	FirstImg, PrevImg, NextImg,	LastImg
		Dim tmpCss : tmpCss = ""
		Dim tmpUrl : tmpUrl = Request.ServerVariables("PATH_INFO")

		If PageCount <> "" Then
		
			FirstImg = "<img src=""/images/common/btn_first.jpg"" alt=""첫페이지"">"
			PrevImg	= "<img src=""/images/common/btn_prev.jpg"" alt=""이전페이지"">"
			NextImg	= "<img src=""/images/common/btn_next.jpg"" alt=""다음페이지"">"
			LastImg	= "<img src=""/images/common/btn_last.jpg"" alt=""마지막페이지"">"
	
				
			StartPage =	Int	( (	Page - 1) /	listPage ) * listPage +	1

			EndPage	= (	StartPage +	listPage ) - 1
			If EndPage > PageCount Then
				EndPage	= PageCount	
			End	If

		
			If Page > 1	Then
				Response.Write "<a href=""JavaScript:goPageAjax(1,'" & tmpUrl & "','#divPopArea .popBox');"" class=""first"">" &	FirstImg & "</a> "
			Else
				Response.Write "<a class=""first"">" & FirstImg & "</a> "
			End	If

			If StartPage > 1 Then
				Response.Write "<a href=""JavaScript:goPageAjax("	& StartPage - 1 & ",'" & tmpUrl & "','#divPopArea .popBox');"" class=""prev"">" & PrevImg & "</a> "
			Else
				Response.Write "<a class=""prev"">" & PrevImg & "</a> "
			End	If
			
			For	i =	StartPage to EndPage
				tmpCss = ""
				
				if i = StartPage Then tmpCss = "num1"
				
				If i = Page	Then
					tmpCss = tmpCss & " on"
					Response.Write "<a class=""" & tmpCss &  """>"	& i	& "</a>	" 
				Else
					Response.Write "<a href=""JavaScript:goPageAjax(" &	i &	",'" & tmpUrl & "','#divPopArea .popBox');"" class=""" & tmpCss & """>" & i & "</a> "
				End	If 
			Next

			If EndPage < PageCount	Then
				Response.Write "<a href=""JavaScript:goPageAjax("	& EndPage + 1 & ",'" & tmpUrl & "','#divPopArea .popBox');"" class=""next"">" & NextImg & "</a> "
			Else
				Response.Write "<a class=""next"">" & NextImg & "</a> "
			End	If
			
			If PageCount > Page Then
				Response.Write "<a href=""JavaScript:goPageAjax(" &	PageCount	& ",'" & tmpUrl & "','#divPopArea .popBox');"" class=""last"">" &	LastImg	& "</a>"
			Else
				Response.Write "<a class=""last"">" & LastImg & "</a>"
			End If
		End If 
  	End	Sub  	

	' ########## 페이징 출력 : 관리자용
  	Sub	SB_PagingWriteAdmin (listPage, RecordCount, PageCount)	
		Dim	StartPage, EndPage,	FirstImg, PrevImg, NextImg,	LastImg
		Dim tmpCss : tmpCss = ""
		
		If PageCount <> "" Then
		
			FirstImg = "<img src=""/admin/images/btn_first.jpg"" alt=""첫페이지"">"
			PrevImg	= "<img src=""/admin/images/btn_prev.jpg"" alt=""이전페이지"">"
			NextImg	= "<img src=""/admin/images/btn_next.jpg"" alt=""다음페이지"">"
			LastImg	= "<img src=""/admin/images/btn_last.jpg"" alt=""마지막페이지"">"
	
				
			StartPage =	Int	( (	Page - 1) /	listPage ) * listPage +	1

			EndPage	= (	StartPage +	listPage ) - 1
			If EndPage > PageCount Then
				EndPage	= PageCount	
			End	If

		
			If Page > 1	Then
				Response.Write "<a href=""JavaScript:goPage(1);"" class=""first"">" &	FirstImg & "</a> "
			Else
				Response.Write "<a class=""first"">" & FirstImg & "</a> "
			End	If

			If StartPage > 1 Then
				Response.Write "<a href=""JavaScript:goPage("	& StartPage - 1 & ");"" class=""prev"">" & PrevImg & "</a> "
			Else
				Response.Write "<a class=""prev"">" & PrevImg & "</a> "
			End	If
			
			For	i =	StartPage to EndPage
				tmpCss = ""
				
				if i = StartPage Then tmpCss = "num1"
				
				If i = Page	Then
					tmpCss = tmpCss & " on"
					Response.Write "<a class=""" & tmpCss &  """>"	& i	& "</a>	" 
				Else
					Response.Write "<a href=""JavaScript:goPage(" &	i &	");"" class=""" & tmpCss & """>" & i & "</a> "
				End	If 
			Next

			If EndPage < PageCount	Then
				Response.Write "<a href=""JavaScript:goPage("	& EndPage + 1 & ");"" class=""next"">" & NextImg & "</a> "
			Else
				Response.Write "<a class=""next"">" & NextImg & "</a> "
			End	If
			
			If PageCount > Page Then
				Response.Write "<a href=""JavaScript:goPage(" &	PageCount	& ");"" class=""last"">" &	LastImg	& "</a>"
			Else
				Response.Write "<a class=""last"">" & LastImg & "</a>"
			End If
		End If 
  	End	Sub  	
%>