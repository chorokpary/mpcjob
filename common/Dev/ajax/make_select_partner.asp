<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : make_select_partner.asp / 파트너 Selectbox
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Parent : Parent = FN_Req("Parent","")
	Dim TargetID : TargetID = FN_Req("TargetID","PartnerDepth2")
	Dim Css : Css = FN_Req("Css","")
	
	If Not IsNumeric(Parent) Then	Parent = 0
	
	Dim ParnterSql : ParnterSql = "Execute uspPartner_Info @Flag = 'CDD2', @Seq='" & Parent & "'" 
	response.write " ": Call SB_MakeHTMLSelect(objDBCon, ParnterSql, TargetID, "", "", "", "[HIDE]", " class=""" & Css & """ ", False)
%>
