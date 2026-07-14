<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : popup.asp / 팝업
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq
	Dim Title, Contents, Height
	
	Call Main()
	
	Sub Main()
		Seq 	= FN_Req("seq","0")
		Call getData()
	End Sub
	
	Sub getData()
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfPopup_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			Set Rs = .Execute
		End With
			
		If Not Rs.Eof Then

			Title 		= Rs("Title")
			Contents 	= Rs("Contents")
			Height		= Rs("Height")
			
			Contents 	= replace(Contents,"target=_parent","target_parent onclick=""javascript:opener.location.href=this.href;self.close();""")
		End If
		Rs.Close
		Set Rs = Nothing
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
	<head>
		<title><%=Title%></title>
		<!--#include virtual="/common/include/head.asp"-->
		<script type="text/javascript">
		<!--
		
		function setCookie( name, value, expiredays )
		{
			var today = new Date();
			today.setDate( today.getDate() + expiredays );
			document.cookie = name + "=" + escape( value ) + "; path=/; expires=" + today.toGMTString() + ";";
		}
		
		function closePopup() 
		{
			setCookie("popup<%=Seq%>", "checked", 1);
			window.close();
		}
		//-->
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
	<div style="height:<%=Height%>px;">
	<%=Contents%>
	</div>
	<div style="width:100%;text-align:right;">
		하루동안 이 창을 열지 않음
		<input type="checkbox" style="margin-right:5px;" onclick="closePopup();">
	</div>
	</body>
</html>
