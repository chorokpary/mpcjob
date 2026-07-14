<%
	Function getIsNew()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspNote_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 		= "ISNEW"
			.Parameters("@RecvSeq").Value 	= Session("ASeq")
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				getIsNew = Rs("Seq")
			Else
				getIsNew = 0
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Function
	
	Dim NewNoteSeq : NewNoteSeq = getIsNew()
	
%>
		<div class="footer">
			<div class="inner"> <address><img src="/admin/images/footer_admin.jpg" alt="서울시 중구 소월로 2길 30 남산트라팰리스 8F (주)한국코퍼레이션 / 대표전화 : 02-3432-6100 / COPYRIGHT (C) HANKOOK CORPORATIONINC. ALL RIGHTS RESERVED" /></address></div>
		</div>
		
		<% If NewNoteSeq > 0 Then  %>
		<script type="text/javascript">
			$(window).load( function(){ openWin("/admin/popup/pop_message.asp?Seq=<%=NewNoteSeq%>","popAlertNote",0,0,450,320,0); });
		</script>
		<% End If %>
