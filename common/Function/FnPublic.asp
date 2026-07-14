<%
'____________________________________________________________________________________
'
' * Discription : FnPublic.asp / 공통 함수.
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2007-12-21 / HYUN / - / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


	' ########## 두개 값을 가져와 일치하면 strRtn 리턴
	Function Fn_SetDefault(getV1, getV2, getRtnT, getRtnF)
	
		If IsNull(getV1) Then	getV1 = ""
		If IsNull(getV2) Then	getV2 = ""

		getV1 = CStr(Trim(getV1))
		getV2 = CStr(Trim(getV2))
		
		If getV1 = getV2 Then 
			Fn_SetDefault = getRtnT
		Else
			Fn_SetDefault = getRtnF
		End If
		
	End Function


	' ########## 배열화된 코드에서 코드 받아 값 출력
	Function Fn_GetArrValue(getArr, getCd)
		Dim fnI, fnV
	
		For fnI = 0 To UBound(getArr,1)
			If UCase(Trim(getArr(fnI,0))) = UCase(Trim(getCd)) Then		fnV = getArr(fnI,1)
		Next
		
		Fn_GetArrValue = fnV
		
	End Function	


	' ########## 두개의 배열 받아서 코드와 네임으로된 배열 반환
	Function Fn_GetArrCD(getArrV, getArrN)
	
		Dim fnRtnArr, fnI

		ReDim fnRtnArr(UBound(getArrV),1)
		
		For fnI = 0 To UBound(getArrV)
			fnRtnArr(fnI,0) 	= getArrV(fnI)
			fnRtnArr(fnI,1) 	= getArrN(fnI)
		Next
	
		Fn_GetArrCD = fnRtnArr
		
	End Function	

	' ########## 문자열을 2차배열로 반환
	Function Fn_SetArr(getStr, getS1, getS2)
		'getS1 : 1차 구분자 / getS2 : 2차 구분자
		'EX : str = Fn_SetArr("1,일/2,이/3,삼","/",",")
	
		Dim fnArr1, fnArr2, fnI, fnJ
		Dim fnLen1, fnLen2
		Dim fnRtnArr
		
		fnArr1 = Split(getStr,getS1)
		fnLen1 = UBound(fnArr1)

		If fnLen1 >= 1 Then
		
			For fnI = 0 To fnLen1
				fnArr2 = Split(fnArr1(fnI),getS2)
				
				If fnLen2 < UBound(fnArr2) Then
					fnLen2 = UBound(fnArr2)
				End If
			Next

			ReDim fnRtnArr(fnLen1,fnLen2)
			
			For fnI = 0 To fnLen1
				fnArr2 = Split(fnArr1(fnI),getS2)

				For fnJ = 0 To UBound(fnArr2)
					fnRtnArr(fnI,fnJ) 	= fnArr2(fnJ)
				Next
			Next
		
		End If
	
		Fn_SetArr = fnRtnArr
		
	End Function	


	' ########## 날짜 포멧 변환
	Function FN_SetDateTimeFormat(strDate,strFormat)
		' strFormat - YYYY:year / MM:month / DD:day / HH: hour / NN:minute / SS : second 
		Dim fnY, fnM, fnD, fnH, fnN, fnS
		Dim fnDate
		
		If IsDate(strDate) Then
			fnY = Right("0000" & Year(strDate),4)
			fnM = Right("00" & Month(strDate),2)
			fnD = Right("00" & Day(strDate),2)

			fnH = Right("00" & Hour(strDate),2)
			fnN = Right("00" & Minute(strDate),2)
			fnS = Right("00" & Second(strDate),2)
			
			fnDate = Replace(strFormat,"YYYY",fnY)
			fnDate = Replace(fnDate,"MM",fnM)
			fnDate = Replace(fnDate,"DD",fnD)
			fnDate = Replace(fnDate,"HH",fnH)
			fnDate = Replace(fnDate,"NN",fnN)
			fnDate = Replace(fnDate,"SS",fnS)
		Else
			fnDate = strDate
		End If
		
		FN_SetDateTimeFormat = fnDate
	End Function

	' ########## 날짜를 타임스탬프로 반환
	Function FN_DatetimeToStamp(argDt)
		Dim fnRtn : fnRtn = ""
		If argDt = "" Or IsNull(argDt) Then
			argDt = now()
		End If
		
		If IsDate(argDt) Then
			fnRtn = DateDiff("s", "01/01/1970 00:00:00", argDt)
		End If
		
		FN_DatetimeToStamp = fnRtn
	End Function

	' ########## 타임스탬프를 날짜로 반환
	Function FN_StampToDate(argStamp)
		Dim fnRtn : fnRtn = ""
		
		If IsNumeric(argStamp) Then
			fnRtn = DateAdd("s", argStamp, "01/01/1970 00:00:00")
		End If
		
		FN_StampToDate = fnRtn
	End Function

	' ########## URL DECODE
	Function FN_URLDecode(Expression)
		Dim strSource, strTemp, strResult, strchr
		Dim lngPos, AddNum, IFKor  
		strSource = Replace(Expression, "+", " ")  
		For lngPos = 1 To Len(strSource)    
			AddNum  = 2
			strTemp = Mid(strSource, lngPos, 1)        
			If strTemp = "%" Then        
				If lngPos + AddNum < Len(strSource) + 1 Then            
					strchr = CInt("&H" & Mid(strSource, lngPos + 1, AddNum))                
					If strchr > 130 Then                 
						AddNum = 5
						IFKor = Mid(strSource, lngPos + 1, AddNum)
						IFKor = Replace(IFKor, "%", "")
						strchr = CInt("&H" & IFKor )                    
					End If                
					strResult = strResult & Chr(strchr)
					lngPos    = lngPos + AddNum                
				End If            
			Else
				strResult = strResult & strTemp
			End If        
		Next
		FN_URLDecode = strResult
	End Function


	' ########## 문자열 복호화
	Public Function Fn_DecryptData(target_data)
		Dim i
		Dim decrypt_result,offset,data_length
		Dim decrypt_result1
		
		decrypt_result = ""
		offset = Asc(mid(target_data, 4, 1))
		data_length = Asc(mid(target_data, 5, 1)) - offset

		For i=0 to data_length-1
		  	offset = Asc(mid(target_data, 2*i+6, 1)) -48
	  	  	decrypt_result = decrypt_result + Chr(Asc(mid(target_data, 2*i+7,1))-offset)
		Next
		
		decrypt_result1=""
		For i=1 To data_length
		 	decrypt_result1=decrypt_result1+chr(asc(mid(decrypt_result,i,1))+1)
		Next
		Fn_DecryptData=decrypt_result1
	End Function
	
	
	' ########## 문자열 암호화
	Public Function Fn_EncryptData(target_data)
		Dim i
		Dim encrypted_result,garbages,offset,data_length
		Dim target_data1
		
		randomize
		encrypted_result = ""
		garbages = 0
		offset = 0

		For i=1 To 3
		 	Do
		 		garbages = Int(Rnd(1)*74)+48
		 	Loop Until ( (garbages > 47 and garbages < 58) or (garbages > 64 and garbages<90) or (garbages>97 and garbages<122) )
			encrypted_result = encrypted_result + Chr(garbages)
		Next

		data_length = len(target_data)
				
		target_data1=""			
			
		For i=1 To data_length
			target_data1=target_data1+chr(asc(mid(target_data,i,1))-1)
		Next

		target_data=target_data1
			
			
		Do
			Do
			 	offset = Int(Rnd(1)*50) + 48
		 	 Loop Until ( (offset > 47 and offset < 58) or (offset > 64 and offset<90) or (offset>97 and offset<122) )
			 garbages = offset + data_length
		Loop Until ( (garbages > 47 and garbages < 58) or (garbages > 64 and garbages<90) or (garbages>97 and garbages<122) )
		
		encrypted_result = encrypted_result + Chr(offset) + Chr(data_length + offset)

		For i=1 To data_length
			Do
				Do
					offset = Int(Rnd(1)*50) + 48
				Loop Until( (offset > 47 and offset < 58) or (offset > 64 and offset<90) or (offset>97 and offset<122) )
				garbages = asc(mid(target_data,i,1)) + (offset - 48)
			Loop Until ( (garbages > 47 and garbages < 58) or (garbages > 64 and garbages<90) or (garbages>97 and garbages<122) )
			
			encrypted_result = encrypted_result + Chr(offset) + Chr(Asc(mid(target_data,i,1))+offset-48)
		Next

		data_length = len(encrypted_result)

		If data_length<50 Then
			For i=1 To 50-data_length
				Do
					garbages = Int(Rnd(1)*74)+48
				Loop Until ( (garbages > 47 and garbages < 58) or (garbages > 64 and garbages<90) or (garbages>97 and garbages<122) )
				
				encrypted_result = encrypted_result + Chr(garbages)
			Next
		End If

		Fn_EncryptData = encrypted_result

	End Function

	'만 나이 계산_20230630
	Function Fn_GetAge(val)
		Dim now_year, now_month, now_day
		Dim user_ymd
		Dim age
		Dim check_month, check_day

		'현재년도
		now_year = Year(Now)

		'생일이 지났는지 여부 체크
		now_month = Month(Now) 
		now_day  = Day(Now)

		user_ymd = Split(val,"-")
		'나이 계산
		age = now_year - user_ymd(0)
		check_month = Int(now_month) - Int(user_ymd(1)) 
		check_day = Int(now_day) - Int(user_ymd(2))
		
		If check_month >= 0 And check_day >= 0 Then
			age = age - 0
		Else 
			age = age - 1
		End If 

		Fn_GetAge= age
	End Function


	' ########## 에러 메시지 출력 후 페이지 이동
	Sub SB_ReturnErr(getMsg, getUrl)
	
		%>
		<script type="text/javascript">
		
			<% If getMsg <> "" Then %>
			alert("<%=getMsg%>");
			<% End If %>
			
			<% If getUrl = "BACK" Then %>
			history.back();
			<% ElseIf getUrl = "CLOSE" Then %>
			self.close();
			<% ElseIf getUrl = "CLOSE_LAYER" Then %>
			closeLayer();
			<% ElseIf getUrl = "SELF" Then%>
			<% Else %>
			window.location.href = "<%=getUrl%>";
			<% End If %>
			
		</script>
		<%
		
		Response.End
		
	End Sub	


	' ########## 에러 메시지 출력 - AJAX
	Sub SB_ReturnErrAjax(getMsg)
		Response.Write("<img src=""/images/common/blank.gif"" onload=""RtnMsg('" & getMsg & "')"">")
		Response.End
	End Sub	


	' ########## DB 연동 SELECT BOX 출력
	Sub SB_MakeHTMLSelect(argObjDB, strSQL, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode)
	
		Dim fnRs
		
		Dim OptionData, OptionCheck, OptionRows
		Dim f
		
		set fnRs = argObjDB.Execute(strSQL)
	
		If fnRs.eof Then
			OptionCheck = False
		Else
			OptionCheck = True
			OptionData = fnRs.getrows
			OptionRows = ubound(OptionData,2)
		End if
			
		fnRs.close
		set fnRs=Nothing
		
		argScript = FN_StrInj(argScript)
		
		If OptionCheck Then
			response.write "<select id=""" & argName & """ name=""" & argName & """ " & argScript & ">" & chr(13)
			
			If argDefaultOpt <> "" Then
				response.write "<option value='" & argDefaultVal & "'>" & argDefaultOpt & "</option>" & chr(13)        
			End If
			
			For f=0 to OptionRows
		  
				response.write "<option value='" 
				If argIsUrlEncode and OptionData(0,f) <> "" Then
					response.write Server.URLEncode(OptionData(0,f))
				Else
					response.write OptionData(0,f)
				End If
				response.write "'"
				If Trim(OptionData(0,f)) = Trim(argSelectVal) Then
					response.write " selected"
				End If
				response.write ">"& OptionData(1,f) &"</option>" & chr(13)
	
			Next
			
			response.write "</select>"        
		Else
			If argNullMsg = "" Then
				response.write "<select id=""" & argName & """ name=""" & argName & """ " & argScript & ">" & chr(13)
				
				If argDefaultOpt <> "" Then
					response.write "<option value='" & argDefaultVal & "'>" & argDefaultOpt & "</option>" & chr(13)        
				End If
				
				response.write "</select>"        
			ElseIf argNullMsg = "[HIDE]" Then
			Else
				response.write argNullMsg
			End If
		End If
	
	End Sub


	' ########## DB 연동 SELECT BOX 출력 : 코드 테이블 연동 > 사용여부 플래그가 True인 경우만
	Sub SB_MakeHTMLSelect_CdTbl(argObjDB, argTbl, argFld, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode)
		Call SB_MakeHTMLSelect_CdTbl_Full(argObjDB, argTbl, argFld, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode, "","USE")
	End Sub

	' ########## DB 연동 SELECT BOX 출력 : 코드 테이블 연동 > 2Depth 일 경우
	Sub SB_MakeHTMLSelect_CdTbl_Dep(argObjDB, argTbl, argFld, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode, argParent)
		Call SB_MakeHTMLSelect_CdTbl_Full(argObjDB, argTbl, argFld, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode, argParent, "USE")
	End Sub


	' ########## DB 연동 SELECT BOX 출력 : 코드 테이블 연동 > 2Depth 일 경우,  사용여부 플래그 구분 가능 [USE, NOT, ALL]
	Sub SB_MakeHTMLSelect_CdTbl_Full(argObjDB, argTbl, argFld, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode, argParent, argIsUse)
	
		Dim fnRs, fnCmd
		
		Dim OptionData, OptionCheck, OptionRows
		Dim f

		argScript = FN_StrInj(argScript)
		
		'set fnRs = argObjDB.Execute(strSQL)
		With objCmd
			.ActiveConnection = argObjDB
			.CommandType = 4
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "SEL"
			.Parameters("@Tbl").Value 	= argTbl
			.Parameters("@Field").Value 	= argFld
			.Parameters("@UseFlag").Value 	= argIsUse
			.Parameters("@Parent").Value 	= argParent		'2Depth의 parent. 없을 경우 ''
			
			' ## 코드 정보
			Set fnRs = .Execute


			If fnRs.eof Then
				OptionCheck = False
			Else
				OptionCheck = True
				OptionData = fnRs.getrows
				OptionRows = ubound(OptionData,2)
			End if

			fnRs.Close
			Set fnRs = Nothing
		End With
	
		
		if OptionCheck Then
			response.write "<select id=""" & argName & """  name=""" & argName & """" & argScript & ">" & chr(13)
			
			If argDefaultOpt <> "" Then
				response.write "<option value=""" & argDefaultVal & """>" & argDefaultOpt & "</option>" & chr(13)        
			End If
			
			For f=0 to OptionRows
		  
				response.write "<option value=""" 
				If argIsUrlEncode and OptionData(0,f) <> "" Then
					response.write Server.URLEncode(FN_InputVal(OptionData(0,f)))
				Else
					response.write FN_InputVal(OptionData(0,f))
				End If
				response.write """"
				If Trim(OptionData(0,f)) = Trim(argSelectVal) Then
					response.write " selected"
				End If
				response.write ">"& OptionData(1,f) &"</option>" & chr(13)
	
			Next
			
			response.write "</select>"        
		Else
			response.write argNullMsg
		End If
	
	End Sub
	
%>