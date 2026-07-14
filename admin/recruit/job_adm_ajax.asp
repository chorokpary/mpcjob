<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_adm_ajax.asp / 채용공고 심사 담당자 목록 ajax
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, RecrSeq, AdmSeq, NewSeq, SearchKey
	Dim kFields, arrData, arrDataLen, i
	Dim StrObjID : StrObjID = "SelAdmSeq"
	
	Call Main()
	
	Sub Main()

		Flag = FN_Req("Flag","INIT")
		RecrSeq = FN_Req("RecrSeq","0")
		AdmSeq = FN_Req("AdmSeq","")
		NewSeq = FN_Req("NewSeq","")
		SearchKey = FN_Req("SearchKey","")
		
		Call getData()
		
		If Flag = "SEARCH" Then
			StrObjID = "SearchAdmSeq"
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		Dim TargetData
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = "uspRecruitCharge_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= Flag
			.Parameters("@RecrSeq").Value 	= RecrSeq
			.Parameters("@AdmSeq").Value 	= AdmSeq
			.Parameters("@NewSeq").Value 	= NewSeq
			.Parameters("@SearchKey").Value = SearchKey
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				' AdmSeq(0), AdmName(1)
				kFields 	= Array("AdmSeq","AdmName")
				arrData = Rs.GetRows(,,kFields)
				arrDataLen 	= UBound(arrData, 2)
			Else
				arrDataLen = -1
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<p>선택된 인원명 : <%=arrDataLen + 1%>명</p>
<div class="innerBox">
	<% For i = 0 To arrDataLen %>
	<p><input type="checkbox" name="<%=StrObjID%>" id="<%=StrObjID%>_<%=arrData(0,i)%>" value="<%=arrData(0,i)%>" /><label for="<%=StrObjID%>_<%=arrData(0,i)%>"><%=arrData(1,i)%></label></p>
	<% Next %>
</div>
