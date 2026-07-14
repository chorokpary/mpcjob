<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : make_select_project.asp / 프로젝트 Selectbox
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2018-03-14 / 이길홍 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim ChargeCode : ChargeCode = FN_Req("ChargeCode","")
	Dim TargetID : TargetID = FN_Req("TargetID","AdminList_1")
	Dim Css : Css = FN_Req("Css","")
	
	If Not IsNumeric(ChargeCode) Then	ChargeCode = 0
	
	Dim RecruitSql : RecruitSql = "Execute uspRecruit_Info2 @Flag = 'ASLST2', @ChargeCode ='"&ChargeCode&"' " 
	response.write " ": Call SB_MakeHTMLSelect(objDBCon, RecruitSql, TargetID, "", "", "", "[HIDE]", " class=""" & Css & """ ", False)
%>
