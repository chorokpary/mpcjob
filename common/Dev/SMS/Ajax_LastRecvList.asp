<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
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
	Dim arrData, arrDataNum, kData, kFields, i, j, k
	Dim MaxLen : MaxLen = 10
	Dim arrNum

	'암호화 
	Dim KEY_PATH, oSeed, returnMsg 

	Call Main()
	
	Sub Main()
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspSmsHistory_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "SMS_LAST"
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' Seq(0), RecvNum(1)

			kFields 	= Array("Seq","RecvNum")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<ul class="person w80" style="left:158px;">
<% 
k = 0
If kData Then 

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)

	For i = 0 To arrDataNum 
		arrNum = Split(arrData(1,i),",")
		If IsArray(arrNum) Then
			If k < MaxLen Then
%><li><a href="#" onclick="SetNum($(this).html());"><%=arrData(1,i)%></a></li><% 
				k = k+1
			End If
		Else 
			For j = 0 To UBound(arrNum)
				If k < MaxLen Then
%><li><a href="#" onclick="SetNum($(this).html());"><%=arrNum(j)%></a></li><% 
					k = k+1
				End If
			Next
		End If
	Next
End If 
%>
</ul>


