<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : Ajax_LastRecvList.asp / SMS 최근 발송 내역
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-30 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim arrData, arrDataNum, kData, kFields, i
	Dim arrNum
	Dim SearchKey

	Call Main()
	
	Sub Main()
		SearchKey = FN_Req("SearchKey","")
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "SMS"
			.Parameters("@UsrName").Value 	= SearchKey
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' UsrSeq(0), UsrName(1), Mobile(2)

			kFields 	= Array("UsrSeq","UsrName","Mobile")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<ul class="person w150">
<% 
If kData Then 
	For i = 0 To arrDataNum 
	%><li><a href="#" onclick="SetNum($(this).html());"><%=arrData(2,i)%>, <%=arrData(1,i)%></a></li><% 
	Next
End If 
%>
</ul>


