<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : make_select_cd_2depth.asp / 코드 테이블 연동 2Depth Selectbox
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag : Flag = FN_Req("Flag","")
	Dim Parent : Parent = FN_Req("Parent","")
	Dim TargetID : TargetID = FN_Req("TargetID","")
	Dim Def : Def = FN_Req("Def","")
	Dim Css : Css = FN_Req("Css","")
	
	Dim Tbl, Field
	
	Select Case Flag 
		Case "RecrPart" : Tbl = "TBL_RECRUIT" : Field = "RecrPart2"
	End Select

	If Parent <> "" Then
		Call SB_MakeHTMLSelect_CdTbl_Dep(objDBCon, Tbl, Field, TargetID, "", "", Def, "", " class=""" & Css & """ ", False, Parent)
	End If
%>
