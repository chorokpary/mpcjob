<%
'____________________________________________________________________________________
'
' * Discription : FnString.asp / 문자열 관련 함수
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2007-12-21 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________

	' ########## Request Injection 처리 (Form 을 우선 순위로 / 업로드 컴포넌트에서는 폼을 우선순위로 할경우 에러 발생)
	Function FN_ReqF(getParam, strDef)
	
		Dim fnStr
		
		' 폼값을 우선순위로..
		If Request.Form(getParam)<> "" Then
			fnStr = Request.Form(getParam)
		Else
			fnStr = Request.QueryString(getParam)
		End If
		
		fnStr = Trim(fnStr)
		fnStr = FN_StrInj(fnStr)
		
		' +++ Default 문자 셋팅
		fnStr = FN_setBlank(fnStr,strDef)
		
		FN_ReqF = fnStr
		
	End Function


	' ########## Request Injection 처리 (QueryString 을 우선 순위로)
	Function FN_Req(getParam, strDef)
	
		Dim fnStr
		
		fnStr = Request(getParam)
		
		fnStr = Trim(fnStr)
		fnStr = FN_StrInj(fnStr)
		
		' +++ Default 문자 셋팅
		fnStr = FN_setBlank(fnStr,strDef)
		
		FN_Req = fnStr
		
	End Function


	' ########## Request Injection & Tag 처리
	Function FN_ReqNoTag(getParam)
	
		Dim fnStr
		
		fnStr = Request(getParam)
		fnStr = Trim(fnStr)
		fnStr = FN_StrInj(fnStr)
		
		fnStr = Replace(fnStr,"<","&lt;")
		fnStr = Replace(fnStr,">","&gt;")
		
		FN_ReqNoTag = fnStr
		
	End Function


	' ########## 문자열 Injection 처리
	Function FN_StrInj(getStr)
	
		Dim fnStr
		
		fnStr = getStr
		fnStr = Trim(fnStr)
		'fnStr = replace(replace(replace(replace(fnStr,"select",""),"update",""),"delete",""),"drop","")

		fnStr= FN_EregiReplace("<(style|script|title|iframe|frmae|ex-pression)(.*)</(style|script|title|iframe|frmae|ex-pression)>","",fnStr)
		fnStr= FN_EregiReplace("<(style|script|title|iframe|frmae|ex-pression)(.*)>","",fnStr)
		fnStr= FN_EregiReplace("<[/]*(style|script|title|iframe|frmae|ex-pression|xmp)>","",fnStr)
		fnStr = Replace(fnStr,"<%","&lt;%")
		'fnStr = Replace(fnStr,"javascript","x-script")
		fnStr = Replace(fnStr,"script","x-script")
		fnStr = Replace(fnStr,"onload","xload")

		FN_StrInj = fnStr
		
	End Function

	
	' ########## 정규표현식을 이용한 문자 대치
	Function FN_EregiReplace(pattern, replace, text)

		Dim fnEreObj : Set fnEreObj= New RegExp
	
		fnEreObj.Pattern= pattern  ' 패턴 설정
		fnEreObj.IgnoreCase = True ' 대소문자 구분 여부
		fnEreObj.Global= True      ' 전체 문서에서 검색
	
		FN_EregiReplace = fnEreObj.Replace(text, replace) ' Replace String

	End Function 


	' ########## NULL 문자 치환
	Function FN_setNull(str,strRtn)
		If IsNull(str) Then
			FN_setNull = strRtn
		Else
			FN_setNull = str
		End	If
	End	Function
	

	' ########## Blank 치환
	Function FN_setBlank(str, strRtn)
		If FN_isBlank(str) Then
			FN_setBlank = strRtn
		Else
			FN_setBlank = str
		End	If
	End	Function
	

	' ########## Blank 체크
	Function FN_isBlank(str)
		If IsNull(str) Or str = "" Then
			FN_isBlank = True
		Else
			FN_isBlank = False
		End	If
	End	Function
	

	' ########## 문자 자르기 (Byte 단위)
	Function FN_StringCut(argStr, argStart, argEnd, argEndStr)
	
		Dim fnI, fnSumByte, fnOneText, fnRtn 
		
		For fnI = 1 To Len(argStr) 
			fnOneText = MID(argStr, fnI, 1) 
			fnSumByte = fnSumByte + FN_StringByte(fnOneText)
			
			If fnSumByte > argEnd Then
				' 지정 바이트 보다 크므로 무조건 1자 삭제
				fnRtn = MID(fnRtn,1,Len(fnRtn)-1) & argEndStr : Exit For 
			Else 
				If argStart <= fnSumByte Then
					fnRtn = fnRtn + fnOneText 
				End If
			End If
		Next
		 
		FN_StringCut = fnRtn 
		
	End Function
	
	' ########## 문자 Byte 리턴
	Function FN_StringByte(argStr)
	
		Dim fnI, fnSumByte, fnOneText 
		
		For fnI = 1 To Len(argStr) 
			fnOneText = MID(argStr, fnI, 1) 
			
			' Windows 2012 에서 한글 ASC 1 반환
			If ASC(fnOneText)=0 Or ASC(fnOneText)=1 Or ASC(fnOneText)=13 Then
				fnSumByte = fnSumByte + 2 
			Else 
				fnSumByte = fnSumByte + 1 
			End If
		Next
		
		FN_StringByte = fnSumByte 
		
	End Function
	


	' ########## HTML 태그 삭제
	Function FN_ClearTag(t)
		Dim tmpStr, strPattern, objRegExp, Matches, Match
	
		tmpStr = t
		
		'+++ 태그 처리 START
		
		strPattern  = "<(.*?)>"
		Set objRegExp           = New RegExp
		objRegExp.Pattern       = strPattern
		objRegExp.Global        = True
		objRegExp.IgnoreCase    = True
		
		Set Matches  = objRegExp.Execute(tmpStr)
	
		For Each Match in Matches   ' Iterate Matches collection.
			tmpStr = replace(tmpStr, Match.Value ,"")
		Next
		
		Set objRegExp   = Nothing
		
		'+++ 태그 처리 END
		
		FN_ClearTag = tmpStr
		
	End Function	

	Function FN_NLToBr(str)
		Dim fnStr
		If Not FN_isBlank(str) Then
			fnStr = FN_ClearTag(str)
			fnStr = Replace(fnStr,chr(13) & chr(10),"<br>")
			fnStr = Replace(fnStr," ","&nbsp;")
		End If
		FN_NLToBr = fnStr
	End Function

	Function FN_BrToNL(str)
		Dim fnStr
		If Not FN_isBlank(str) Then
			fnStr = str
			fnStr = Replace(fnStr,"<br>",chr(13) & chr(10))
			fnStr = Replace(fnStr,"&nbsp;"," ")
		End If
		FN_BrToNL = fnStr
	End Function

	Function HTMLDecode(sText)
	    Dim I
	    sText = Replace(sText, "&quot;", Chr(34))
	    sText = Replace(sText, "&lt;"  , Chr(60))
	    sText = Replace(sText, "&gt;"  , Chr(62))
	    'sText = Replace(sText, "&nbsp;", Chr(32))
	    sText = Replace(sText, "&amp;" , Chr(38))
	    For I = 1 to 255
	        sText = Replace(sText, "&#" & I & ";", Chr(I))
	    Next
	    HTMLDecode = sText
	End Function	

	Function HTMLEncode(sText)
	    Dim I
	    sText = Replace(sText, Chr(38), "&amp;")
	    sText = Replace(sText, Chr(34), "&quot;")
	    sText = Replace(sText, Chr(60), "&lt;")
	    sText = Replace(sText, Chr(62), "&gt;")
	    'sText = Replace(sText, Chr(32), "&nbsp;")
	    HTMLEncode = sText
	End Function	

	Function FN_InputVal(sText)
		sText = FN_setBlank(sText,"")
		sText = Replace(sText,"""","&quot;")
	    FN_InputVal = sText
	End Function	

%>