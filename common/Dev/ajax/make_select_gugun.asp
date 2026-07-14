<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : make_select_gugun.asp / 구군 Selectbox
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim SIDO : SIDO = FN_Req("SIDO","")
	Dim TargetID : TargetID = FN_Req("TargetID","sGugun")
	Dim Css : Css = FN_Req("Css","")
	
	Dim GugunSql : GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '" & SIDO & "' GROUP BY GUGUN ORDER BY GUGUN" 
	'response.write "" : 
	Call SB_MakeHTMLSelect(objDBCon, GugunSql, TargetID, "지역중분류", "", "", "", " class=""" & Css & """ ", False)
%>
