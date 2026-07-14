<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Dev/MsDao.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_list.asp / 회원 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-10 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2014-01-07 / 강미현 / 4M / 희망분야 등록항목 삭제
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim sGender, sHope, sHope2, sStatus, sSido, sGugun, sBDate, sEDate, sBlack, sSearch, sAge, sAge2, sBirth, sBirth2
	Dim sKey, sWord, sSearchHistory, sSearchIn
	Dim arrSearch, arrSearchData, strSearchInfor, strSearchKey
	Dim SortKey, SortDIr
	Dim StrCntToday : StrCntToday = 0
	Dim strFilter

	Call Main()

	Sub Main()

		If Request.Form = "" Then
			strFilter = Replace(Request.QueryString,"&","[^]")
		Else
			strFilter = Replace(Request.Form,"&","[^]")
		End If
		
		
		listPage 	= 10
		listRecord 	= FN_ReqF("ListSize","50")	'목록갯수
		Page 		= FN_ReqF("Page","1")

		sKey 	= FN_ReqF("sKey","MOBILE")
		sWord 	= FN_InputVal(FN_ReqF("sWord",""))
		sSearchIn = FN_ReqF("sSearchIn","")

		sGender = FN_ReqF("sGender","")
		sHope 	= FN_ReqF("sHope","")
		sHope2 	= FN_ReqF("sHope2","")
		sStatus = FN_ReqF("sStatus","")
		sSido 	= FN_ReqF("sSido","")
		sGugun 	= FN_ReqF("sGugun","")
		sBDate 	= FN_ReqF("sBDate","")
		sEDate 	= FN_ReqF("sEDate","")

		sBlack 	= FN_ReqF("sBlack","")
		sAge 	= FN_ReqF("sAge","")
		sAge2 	= FN_ReqF("sAge2","")

		SortKey = FN_ReqF("SortKey","")
		SortDIr = FN_ReqF("SortDIr","DESC")

		If sAge = "" Then 
			sAge = "999"
			sAge2 = "999"
			sBirth = ""
		Else
			sBirth = Year(Date) - sAge + 1
			sBirth2 = Year(Date) - sAge2 + 1
		End If 

		If Not IsDate(sBDate) Then	sBDate = ""
		If Not IsDate(sEDate) Then	sEDate = ""
		If sSearchIn = "Y" Then
			sSearchHistory = FN_ReqF("sSearchHistory","")
			sSearchHistory = sSearchHistory & Chr(13) & Chr(10) & sKey & Chr(9) & sWord
		Else
			sSearchHistory = sKey & Chr(9) & sWord
		End If
		If sSearchHistory <> "" Then
			arrSearch = Split(sSearchHistory,Chr(13) & Chr(10))
			ReDim arrSearchData(UBound(arrSearch))
			For i = 0 To UBound(arrSearch)
				arrSearchData(i) = Split(arrSearch(i),Chr(9))
				
				Select Case Trim(arrSearchData(i)(0))
					Case "ID" : strSearchKey = "아이디 "
					Case "NAME" : strSearchKey = "이름"
					Case "ADDR" : strSearchKey = " 거주지 "
					Case "MOBILE" : strSearchKey = "핸드폰 "
					Case "PROJECT" : strSearchKey = "최근 지원한 프로젝트"
					Case "CAREER" : strSearchKey = "경력"
				End Select 
				
				If arrSearchData(i)(1) <> "" Then
					strSearchInfor = strSearchInfor & "<span><strong>+ " & strSearchKey & " : </strong>" & arrSearchData(i)(1) & "</span>"
				End If
			Next
		Else
			strSearchInfor = ""
		End If
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 20
		End If
		
		' 결과내 재검색 X , SP 사용
		'Call getList()
		'결과내 재검색, Query 사용
		Call getList2()
		
	End Sub
	
	' 결과 내 재검색 (Query)
	Sub getList2()
		Dim Rs, Sql, Sql_Where
		Dim Sql_Order1, Sql_Order2, Sql_Order3
		Dim listFirst, listSize, TotalPage
		Dim n
		
		'// ##### Today's Registration
		Sql = "SELECT COUNT(*) As CntToday FROM TBL_USER WITH (nolock) WHERE CONVERT(nchar(10),RegDate,21) = CONVERT(nchar(10),GETDATE(),21) AND IsOut = 0 AND UsrSeq !=0;"
		
		'// ##### 목록 검색 (TotalCnt , List 공통사용)
		If sGender <> "" Then 	Sql_Where = " AND D.Gender = ? "
		If sHope <> "" Then 	Sql_Where = Sql_Where & " AND R.RecrPart = ? "
		If sHope2 <> "" Then 	Sql_Where = Sql_Where & " AND R.RecrPart2 = ? "
		If sStatus <> "" Then 	Sql_Where = Sql_Where & " AND AI.Status = ? "
		If sBDate <> "" Then 	Sql_Where = Sql_Where & " AND CONVERT(nvarchar(10),D.ModDate,21) >= ? "
		If sEDate <> "" Then 	Sql_Where = Sql_Where & " AND CONVERT(nvarchar(10),D.ModDate,21) <= ? "
		If sSido <> "" Then 	Sql_Where = Sql_Where & " AND D.AddrSido = ? "
		If sGugun <> "" Then 	Sql_Where = Sql_Where & " AND D.AddrGugun = ? "
		If sBlack = "1" Then 	
			Sql_Where = Sql_Where & " AND D.IsBlack = '1' "
		ElseIf sBlack = "2" Then 
			Sql_Where = Sql_Where & " AND D.IsSearch = '1'"
		ElseIf sBlack = "3" Then 
			Sql_Where = Sql_Where & " AND D.IsBlack = '0' AND D.IsSearch ='0'"
		Else 
		End If 
		If sBirth <> "" Then 	Sql_Where = Sql_Where & " AND (LEFT(D.Birthday,4) >= '"&sBirth2&"' AND LEFT(D.Birthday,4) <= '"&sBirth&"')"

		If Session("ALevel") = 2 Then 
			'Sql_Where = Sql_Where & " AND D.UsrSeq in (SELECT distinct(UsrSeq) FROM TBL_APPLY WHERE recrseq in ( SELECT RecrSeq FROM TBL_RECRUIT_CHARGE WITH (nolock) WHERE AdmSeq = '"&Session("ASeq")&"')) "
			Sql_Where = Sql_Where & " AND D.UsrSeq in (SELECT distinct(UsrSeq) FROM TBL_APPLY WHERE recrseq in ( SELECT RecrSeq FROM TBL_RECRUIT WITH (nolock) WHERE ChargeCode = '"&Session("ASeq")&"')) "
		End If 

		If sSearchHistory <> "" Then
			For n = 0 To UBound(arrSearch)
				Select Case Trim(arrSearchData(n)(0))
					Case "ID" : Sql_Where = Sql_Where & " AND D.UsrID LIKE '%' + ? + '%' "
					Case "NAME" : Sql_Where = Sql_Where & " AND D.UsrName LIKE '%' + ? + '%' "
					Case "ADDR" : Sql_Where = Sql_Where & " AND (D.AddrSido LIKE '%' + ? + '%' OR D.AddrGugun LIKE '%' + ? + '%') "
					Case "MOBILE" : Sql_Where = Sql_Where & " AND REPLACE(D.Mobile,'-','') LIKE '%' + ? + '%' "
					Case "PROJECT" : Sql_Where = Sql_Where & " AND P.PrjName LIKE '%' + ? + '%' "
					Case "CAREER" : Sql_Where = Sql_Where & " AND D.UsrSeq IN (SELECT UsrSeq FROM TBL_USER_CAREER_EZ (nolock) WHERE  Career LIKE '%' + ? + '%') "
				End Select 
			Next
		End If
		
		'// ##### TotalCnt
		Sql = Sql & 	" SELECT COUNT(*) AS TotalCnt, CEILING(CAST(COUNT(*) AS FLOAT) / "& listRecord &") AS TotalPage " &_
					" FROM TBL_USER AS D (nolock)" &_
					" LEFT OUTER JOIN ( " &_
					" 	SELECT MAX(AppSeq) AS AppSeq, UsrSeq " &_
					" 	FROM TBL_APPLY WITH (nolock) " &_
					" 	WHERE RecrSeq > 0 " &_
					" 	GROUP BY UsrSeq " &_
					" ) AS A " &_
					" ON D.UsrSeq = A.UsrSeq " &_
					" LEFT OUTER JOIN TBL_APPLY AS AI WITH (nolock) " &_
					" ON A.AppSeq = AI.AppSeq " &_
					" LEFT OUTER JOIN TBL_RECRUIT AS R WITH (nolock) " &_
					" ON AI.RecrSeq = R.RecrSeq " &_
					" LEFT OUTER JOIN TBL_PROJECT AS P WITH (nolock) " &_
					" ON P.PrjSeq = R.PrjSeq	" &_
					" WHERE D.IsOut = 0 AND D.UsrSeq !=0 " & Sql_Where & " ;"
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 1 ' adCmdText
			.CommandText =Sql
'			.Parameters.Refresh
			
			If sGender <> "" Then	.Parameters.Append .CreateParameter("sGender", adWChar, adParamInput, 1, sGender)
			If sHope <> "" Then	.Parameters.Append .CreateParameter("sHope", adVarWChar, adParamInput, 100, sHope)
			If sHope2 <> "" Then	.Parameters.Append .CreateParameter("sHope2", adVarWChar, adParamInput, 100, sHope2)
			If sStatus <> "" Then	.Parameters.Append .CreateParameter("sStatus", adVarWChar, adParamInput, 2, sStatus)
			If sBDate <> "" Then	.Parameters.Append .CreateParameter("sBDate", adVarWChar, adParamInput, 10, sBDate)
			If sEDate <> "" Then	.Parameters.Append .CreateParameter("sEDate", adVarWChar, adParamInput, 10, sEDate)
			If sSido <> "" Then		.Parameters.Append .CreateParameter("sSido", adVarWChar, adParamInput, 8, sSido)
			If sGugun <> "" Then	.Parameters.Append .CreateParameter("sGugun", adVarWChar, adParamInput, 30, sGugun)
			'If sBlack <> "" Then	.Parameters.Append .CreateParameter("sBlack", adVarWChar, adParamInput, 30, sBlack)
			'If sBirth <> "" Then	.Parameters.Append .CreateParameter("sBirth", adVarWChar, adParamInput, 30, sBirth)
			If sSearchHistory <> "" Then
				For n = 0 To UBound(arrSearch)
					Select Case Trim(arrSearchData(n)(0))
						Case "ID", "NAME", "PROJECT", "CAREER"
							.Parameters.Append .CreateParameter("sWord", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
						Case "MOBILE"
							.Parameters.Append .CreateParameter("sWord", adVarWChar, adParamInput, 200, Replace(Replace(Trim(arrSearchData(n)(1)),"[","[[]"),"-",""))
						Case "ADDR" : 
							.Parameters.Append .CreateParameter("sWord1", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
							.Parameters.Append .CreateParameter("sWord2", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
					End Select 
				Next
			End If
			
			Set Rs = .Execute
			
		End With
		
		If Not Rs.Eof Then
			StrCntToday = Rs("CntToday")
		End If
		
		Set Rs = Rs.NextRecordSet()
			
		If Not Rs.Eof Then
			RecordCount = Rs("TotalCnt")
			TotalPage = Rs("TotalPage")
		End If

	

		Rs.Close
		Set Rs = Nothing
		
		
		'// ##### List Info
		If TotalPage < 1 Then TotalPage = 1
		If Page > TotalPage Then Page = TotalPage
		If Page < TotalPage Then
			listSize = listRecord
		Else
			listSize = RecordCount - ((Page - 1) * listRecord)
		End If

		If SortKey = "GENDER" Then
			Sql_Order1 = " D.Gender [ORDER],"
			Sql_Order2 = " Gender [ORDER],"
			Sql_Order3 = " Gender [ORDER],"
		ElseIf SortKey = "AGE" Then
			Sql_Order1 = " CASE ISNUMERIC(LEFT(D.BirthDay,4)) WHEN 1 THEN  YEAR(GETDATE())-LEFT(D.BirthDay,4) + 1 END [ORDER],"
			Sql_Order2 = " CASE ISNUMERIC(LEFT(BirthDay,4)) WHEN 1 THEN  YEAR(GETDATE())-LEFT(BirthDay,4) + 1 END [ORDER],"
			Sql_Order3 = " CASE ISNUMERIC(LEFT(BirthDay,4)) WHEN 1 THEN  YEAR(GETDATE())-LEFT(BirthDay,4) + 1 END [ORDER],"
		End If
		
		If SortDir = "DESC" Then
			Sql_Order1 = Replace(Sql_Order1,"[ORDER]","DESC")
			Sql_Order2 = Replace(Sql_Order2,"[ORDER]","ASC")
			Sql_Order3 = Replace(Sql_Order3,"[ORDER]","DESC")
		Else
			Sql_Order1 = Replace(Sql_Order1,"[ORDER]","ASC")
			Sql_Order2 = Replace(Sql_Order2,"[ORDER]","DESC")
			Sql_Order3 = Replace(Sql_Order3,"[ORDER]","ASC")
		End If
		
		Sql = 	" SELECT T2.TotalCnt, T2.UsrSeq, T2.UsrID, T2.UsrName, T2.Gender, T2.Birthday, T2.AddrSido, T2.AddrGugun, T2.Mobile, T2.HopePart, T2.HopePart2, T2.RegDate, T2.ModDate, T2.IsBlack, T2.Status, T2.PartnerDepth1, T2.PartnerDepth2 , T2.PrjName , T2.EzCareer, T2.CntEzCareer, T2.RecrPart, T2.RecrPart2, T2.UsrType, T2.IsSearch " &_
				" 	, ISNULL((SELECT COUNT(*) AS CntMemo FROM TBL_USER_MEMO WITH (nolock) WHERE UsrSeq = T2.UsrSeq),0) AS CntMemo, T2.CntStat  " &_ 	
				" FROM ( " &_
				" 	SELECT TOP "& listSize &" T.TotalCnt, T.UsrSeq, T.UsrID, T.UsrName, T.Gender, T.Birthday, T.AddrSido, T.AddrGugun, T.Mobile, T.HopePart, T.HopePart2, T.RegDate, T.ModDate, T.IsBlack, T.Status, T.PartnerDepth1, T.PartnerDepth2 , T.PrjName , T.EzCareer, T.CntEzCareer, T.RecrPart, T.RecrPart2, T.UsrType, T.IsSearch, T.CntStat " &_
				" 	FROM (" &_
				" 		SELECT TOP " & (Page * listRecord) & " 0 AS TotalCnt, D.UsrSeq " &_
				" 		, D.UsrID, D.UsrName, D.Gender, D.Birthday, D.AddrSido, D.AddrGugun, D.Mobile, D.HopePart, D.HopePart2, D.RegDate, D.ModDate, D.IsBlack, Isnull(D.UsrType,0) as UsrType, D.IsSearch  " &_
				" 		, CC.CdVal AS Status, AI.PartnerDepth1, AI.PartnerDepth2 , P.PrjName " &_
				" 		, EC.EzCareer, EC.CntEzCareer " &_
				" 		, R.RecrPart, R.RecrPart2, A.CntStat " &_
				" 		FROM TBL_USER AS D WITH (nolock) " &_
				" 		LEFT OUTER JOIN ( " &_
				" 			SELECT MAX(AppSeq) AS AppSeq, UsrSeq, (select count(*) from Tbl_apply where status in ('04','06','08') and usrSeq = AP.UsrSeq) as CntStat " &_
				" 			FROM TBL_APPLY as AP WITH (nolock) " &_
				" 			WHERE RecrSeq > 0 " &_
				" 			GROUP BY UsrSeq " &_
				" 		) AS A " &_
				" 		ON D.UsrSeq = A.UsrSeq " &_
				" 		LEFT OUTER JOIN TBL_APPLY AS AI WITH (nolock) " &_
				" 		ON A.AppSeq = AI.AppSeq " &_
				" 		LEFT OUTER JOIN TBL_RECRUIT AS R WITH (nolock) " &_
				" 		ON AI.RecrSeq = R.RecrSeq " &_
				" 		LEFT OUTER JOIN TBL_PROJECT AS P WITH (nolock) " &_
				" 		ON P.PrjSeq = R.PrjSeq " &_
				" 		LEFT OUTER JOIN TBL_CONF_CODE AS CC WITH (nolock)	 " &_
				" 		ON AI.Status = CC.CdKey AND CC.Tbl = 'TBL_APPLY' AND CC.Field = 'Status' " &_
				" 		LEFT OUTER JOIN ( " &_
				" 			SELECT MIN(Career) AS EzCareer, COUNT(*) AS CntEzCareer, UsrSeq " &_
				" 			FROM TBL_USER_CAREER_EZ WITH (nolock) " &_
				" 			GROUP BY UsrSeq " &_
				" 		) AS EC " &_
				" 		ON D.UsrSeq = EC.UsrSeq " &_
				" 		WHERE D.IsOut = 0 AND D.UsrSeq !=0 " & Sql_Where &_
				" 		ORDER BY " & Sql_Order1 & " D.ModDate DESC, D.UsrSeq DESC" &_
				" 	) AS T" &_
				" 	ORDER BY " & Sql_Order2 & " T.ModDate ASC, T.UsrSeq ASC " &_
				" ) AS T2 " &_
				" ORDER BY " & Sql_Order3 & " T2.ModDate DESC, T2.UsrSeq DESC "
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 1 ' adCmdText
			.CommandText =Sql
			'.Parameters.Refresh

			'If sGender <> "" Then	.Parameters.Append .CreateParameter("sGender", adWChar, adParamInput, 1, sGender)
			'If sHope <> "" Then	.Parameters.Append .CreateParameter("sHope", adVarWChar, adParamInput, 100, sHope)
			'If sHope2 <> "" Then	.Parameters.Append .CreateParameter("sHope2", adVarWChar, adParamInput, 100, sHope2)
			'If sStatus <> "" Then	.Parameters.Append .CreateParameter("sStatus", adVarWChar, adParamInput, 2, sStatus)
			'If sBDate <> "" Then	.Parameters.Append .CreateParameter("sBDate", adVarWChar, adParamInput, 10, sBDate)
			'If sEDate <> "" Then	.Parameters.Append .CreateParameter("sEDate", adVarWChar, adParamInput, 10, sEDate)
			'If sSido <> "" Then		.Parameters.Append .CreateParameter("sSido", adVarWChar, adParamInput, 8, sSido)
			'If sGugun <> "" Then	.Parameters.Append .CreateParameter("sGugun", adVarWChar, adParamInput, 30, sGugun)
			'If sSearchHistory <> "" Then
			'	arrSearch = Split(sSearchHistory,Chr(13) & Chr(10))
			'	For n = 0 To UBound(arrSearch)
			'		Select Case Trim(arrSearchData(n)(0))
			'			Case "ID", "NAME", "MOBILE", "PROJECT", "CAREER"
			'				.Parameters.Append .CreateParameter("sWord", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
			'			Case "ADDR" : 
			'				.Parameters.Append .CreateParameter("sWord1", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
			'				.Parameters.Append .CreateParameter("sWord2", adVarWChar, adParamInput, 200, Replace(Trim(arrSearchData(n)(1)),"[","[[]"))
			'		End Select 
			'	Next
			'End If
			
			Set Rs = .Execute

		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0)(: sp와 cols를 맞추기 위해 임의 삽입), UsrSeq(1), UsrID(2), UsrName(3), Gender(4), Birthday(5), AddrSido(6)
			' AddrGugun(7), Mobile(8), HopePart(9), RegDate(10), ModDate(11), IsBlack(12) 
			' CntMemo(13), Status(14), PartnerDepth1(15), PartnerDepth2(16), PrjName(17), EzCareer(18), CntEzCareer(19), HopePart2(20), RecrPart(21), RecrPart2(22), UsrType(23), IsSearch(24), CntStat(25)
			kFields 	= Array("TotalCnt","UsrSeq","UsrID","UsrName","Gender","Birthday","AddrSido","AddrGugun","Mobile","HopePart","RegDate","ModDate","IsBlack","CntMemo","Status","PartnerDepth1","PartnerDepth2","PrjName","EzCareer","CntEzCareer","HopePart2","RecrPart","RecrPart2","UsrType","IsSearch","CntStat")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True

			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing
	End Sub
	
	' 결과 내 재검색 사용 전 (SP)
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4	' adCmdStoredProc
			.CommandText ="uspUser_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= Fn_SetDefault(sKey,"MOBILE",Replace(sWord,"-",""),sWord) 

			.Parameters("@sGender").Value 	= sGender
			.Parameters("@sHope").Value 	= sHope
			.Parameters("@sHope2").Value 	= sHope2
			.Parameters("@sStatus").Value 	= sStatus
			.Parameters("@sSido").Value 	= sSido
			.Parameters("@sGugun").Value 	= sGugun
			.Parameters("@sBDate").Value 	= sBDate
			.Parameters("@sEDate").Value 	= sEDate
			If SortKey <> "" And SortDir = "DESC" Then
				.Parameters("@SortDesc").Value 	= SortKey
			ElseIf SortKey <> "" And SortDir = "ASC" Then
				.Parameters("@SortAsc").Value 	= SortKey
			End If
			
			Set Rs = .Execute
		End With
		
		If Not Rs.Eof Then
			StrCntToday = Rs("CntToday")
		End If
		
		Set Rs = Rs.NextRecordSet()
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), UsrSeq(1), UsrID(2), UsrName(3), Gender(4), Birthday(5), AddrSido(6)
			' AddrGugun(7), Mobile(8), HopePart(9), RegDate(10), ModDate(11), IsBlack(12)
			' CntMemo(13), Status(14), PartnerDepth1(15), PartnerDepth2(16), PrjName(17), EzCareer(18), CntEzCareer(19), HopePart2(20), RecrPart(21), RecrPart2(22), UsrType(23), IsSearch(24)
			kFields 	= Array("TotalCnt","UsrSeq","UsrID","UsrName","Gender","Birthday","AddrSido","AddrGugun","Mobile","HopePart","RegDate","ModDate","IsBlack","IsSearch","CntMemo","Status","PartnerDepth1","PartnerDepth2","PrjName","EzCareer","CntEzCareer","HopePart2","RecrPart","RecrPart2","IsSearch")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True

			RecordCount = arrData(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>회원관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript">
		var HopePart2 = "<%=sHope2%>";

		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원을 삭제하시겠습니까?")){
					$("#Flag").val("DEL");
					$("#frmList").attr({action:"mem_proc.asp", method:"post"}).submit();
				}
			}
		}

		function chkBlack(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원의 블랙리스트 상태를 변경하시겠습니까?")){
					$("#Flag").val("BLK");
					$("#frmList").attr({action:"mem_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function chkSearch(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원의 서칭거부 상태를 변경하시겠습니까?")){
					$("#Flag").val("SEARCH");
					$("#frmList").attr({action:"mem_proc.asp", method:"post"}).submit();
				}
			}
		}

		function chkAgree(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원의 동의 상태를 변경하시겠습니까?")){
					$("#Flag").val("AGREE");
					$("#frmList").attr({action:"mem_proc.asp", method:"post"}).submit();
				}
			}
		}

		function doPrint(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				openWin("mem_print.asp?Seq="+ getCheckedVal("Seq") ,"popPrint",0,0,930,650,1);
			}
		}

		function doOpenSms(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				openWin("/admin/popup/sms.asp?Flag=USER&Seq="+ getCheckedVal("Seq") ,"popSMS",0,0,948,550,1);
			}
		}

		function doDownXls(){
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원을 엑셀로 다운 받으시겠습니까?")){
					$("#frmList").attr({action:"mem_xls.asp", method:"post"}).submit();
				}
			}
		}

		function doRecruit(){
			if (!$("#selPartnerD1").val()){
				alert("채용 Root를 선택해주세요.");
				return;
			}
			if (!$("#RecrSeq").val()){
				alert("프로젝트(채용공고)를 선택해주세요.");
				return;
			}
			
			if (cntChecked("#frmList")==0){
				alert("선택된 회원이 없습니다.");
			}else{
				if(confirm("선택한 회원을 채용공고 등록 하시겠습니까?")){
					$("#Flag").val("RECR");
					$("#frmList").attr({action:"mem_proc.asp", method:"post"}).submit();
				}
			}
		}

		function showKeyword(){
			/*
			if ($("#sKey").val()=="ADDR"){
				$("#sWord").hide();
				$("#sSido").show();
				$("#sGugun").show();
				$("#sWord").val("");
			}else{
				$("#sWord").show();
				$("#sSido").hide();
				$("#sGugun").hide();
			}
			*/
			
			getHopePart2();
		}
		
		function fillKeywordAddr(){
			initPage();
			//$("#sWord").val($("#sSido").val() + " " + $("#sGugun").val());
			if ($(this).attr("id") == "sSido"){
				getAddrGugun($(this).val());
			}
		}
		
		function getAddrGugun(sido){
			var url = "/common/dev/ajax/make_select_gugun.asp";
			var param = "SIDO=" + sido + "&TargetID=sGugun&Css=w110";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divGugun").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function getPartner(){
			var url = "/common/dev/ajax/make_select_partner.asp";
			var param = "Parent=" + $(this).val() + "&TargetID=PartnerDepth2";
			
			$("#PartnerDepth1").val($(this).find("option:selected").text());
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divPartner").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		// 희망분야 2Depth
		function getHopePart2(){
			var url = "/common/dev/ajax/make_select_cd_2depth.asp";
			var param = "Flag=RecrPart&Parent=" + $("#sHope").val() + "&TargetID=sHope2&Def=" + HopePart2 + "&Css=w95";
			HopePart2 = "";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divHopePart2").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
		function goSort(key,dir){
			$("#SortKey").val(key);
			$("#SortDir").val(dir);
			goPage(<%=Page%>);
		}
		
		function goForm(seq){
			//window.location.href="mem_form.asp?Page=<%=Page%>&Seq="+seq;
			window.location.href="mem_form.asp?Seq="+seq + "&Filter=<%=strFilter%>";
		}
	
		function goUserMemo(seq){ 
			openWin("../member/mem_memo.asp?Seq="+ seq,"popMemo",0,0,930,620,0);
		}

		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnDel").click(chkDel);
		    $("#btnBlack").click(chkBlack);
			$("#btnSearch").click(chkSearch);
		    $("#btnSMS").click(doOpenSms);
		    $("#btnPrint").click(doPrint);
			$("#btnAgree").click(chkAgree);
		    
		    $("#select_box > option:selected").val();

		    $("#btnXls").click(doDownXls);
		    $("#btnReg").click(function(){ goForm(0); });
		    
		    $("#btnRecr").click(doRecruit);

		    $("#ListSize").change(function(){ goPage(1); });
		    $("#sGender").change(initPage);
		    $("#sHope").change(initPage,getHopePart2);
		    $("#sHope2").change(initPage);
		    $("#sStatus").change(initPage);
		    $("#sKey").change(showKeyword);
		    $("#sBDate").change(initPage);
		    $("#sEDate").change(initPage);
		    $("#sSido").change(fillKeywordAddr);
		    $("#sGugun").live("change",fillKeywordAddr);

		    $("#selPartnerD1").change(getPartner);
		    

			$("#sBDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#sEDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});

		})
		
		$(window).load(showKeyword);
	</script>
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents2">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>회원관리 ( 오늘 가입한 신규 회원수는 [ <%=StrCntToday%> ]명 입니다. )</h2>
					<form name="frmPage" id="frmPage" method="post" action="mem_list.asp">
					<input type="hidden" id="Page" name="Page">
					<input type="hidden" id="SortKey" name="SortKey" value="<%=SortKey%>">
					<input type="hidden" id="SortDir" name="SortDir" value="<%=SortDir%>">
					<div class="searchArea">
						<div class="search_detail">
							<div class="fl_l">
								<span class="group">
									<span class="tit">성별</span> 
									<select  name="sGender" id="sGender" title="성별검색" class="w65">
										<option value="">성별</option>
										<option value="1"<%=Fn_SetDefault(sGender,"1"," selected","")%>>남</option>
										<option value="2"<%=Fn_SetDefault(sGender,"2"," selected","")%>>여</option>
									</select>
								</span>

								<span class="group">
									<span class="tit">희망분야</span> 
									<%' response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER", "HopePart", "sHope", "희망분야", "", sHope, "", " class=""w95"" ", False)%>
									<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "RecrPart", "sHope", "희망분야", "", sHope, "", " class=""w95"" ", False)%>
									<span id="divHopePart2"></span>
								</span>
									
								<span class="group">
									<span class="tit">진행단계</span> 
									<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_APPLY", "Status", "sStatus", "진행단계", "", sStatus, "", " class=""w105"" ", False)%>
								</span>


								<span class="group">
									<span class="tit">회원상태</span> 
									<select  name="sBlack" id="sBlack" title="회원상태" class="w65">
										<option value="">전체</option>
										<option value="1"<%=Fn_SetDefault(sBlack,"1"," selected","")%>>블랙리스트</option>
										<option value="2"<%=Fn_SetDefault(sBlack,"2"," selected","")%>>서칭거부</option>
										<option value="3"<%=Fn_SetDefault(sBlack,"3"," selected","")%>>일반</option>
									</select>
								</span>
								<span class="group">
									<span class="tit">연령</span> 
									<select  name="sAge" id="sAge" title="연령" class="w65">
										<option value="">전체</option>
										<%For i=0 To 100%>
										<option value="<%=i%>" <%If CInt(sAge) = i Then%>selected<%End If%>><%=i%></option>
										<%Next%>
									</select> ~ 
									<select  name="sAge2" id="sAge2" title="연령" class="w65">
										<option value="">전체</option>
										<%For i=0 To 100%>
										<option value="<%=i%>" <%If CInt(sAge2) = i Then%>selected<%End If%>><%=i%></option>
										<%Next%>
									</select> 
								</span>

							</div>
							
							<div class="fl_r">
								<select name="ListSize" id="ListSize" title="출력갯수" style="width:65px;">
									<option value="20">출력개수</option>
									<option value="5"<%=Fn_SetDefault(listRecord,"5"," selected","")%>>5개</option>
									<option value="10"<%=Fn_SetDefault(listRecord,"10"," selected","")%>>10개</option>
									<option value="20"<%=Fn_SetDefault(listRecord,"20"," selected","")%>>20개</option>
									<option value="30"<%=Fn_SetDefault(listRecord,"30"," selected","")%>>30개</option>
									<option value="50"<%=Fn_SetDefault(listRecord,"50"," selected","")%>>50개</option>
									<option value="100"<%=Fn_SetDefault(listRecord,"100"," selected","")%>>100개</option>
									<option value="300"<%=Fn_SetDefault(listRecord,"300"," selected","")%>>300개</option>
									<option value="500"<%=Fn_SetDefault(listRecord,"500"," selected","")%>>500개</option>
								</select>
							</div>
						</div>
						<div class="search_detail pt10">
							<div class="fl_l">
								<span class="group">
									<span class="tit">수정일</span> 
									<input type="text" id="sBDate" name="sBDate" value="<%=sBDate%>" style="margin-right:5px; width:65px;" maxlength="10" /> ~ 
									<input type="text" id="sEDate" name="sEDate" value="<%=sEDate%>" style="margin-right:5px; width:65px;" maxlength="10" />
								</span>
								
								<span class="group">
									<span class="tit">거주지</span> 
									<%
									Dim SidoSql : SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
									response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "sSido", "지역", "", sSido, "", " class=""w80"" ", False)
									%>
									<span id="divGugun">
									<%
									Dim GugunSql : GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '" & sSido & "' GROUP BY GUGUN ORDER BY GUGUN" 
									response.write " ": Call SB_MakeHTMLSelect(objDBCon, GugunSql, "sGugun", "지역중분류", "", sGugun, "", " class=""w110"" ", False)
									%>
									</span>
								</span>
							</div>
							
						</div>
						<div class="search_date">
							<div class="fl_l">

								<span class="tit_date">키워드 검색</span> 
								<select name="sKey" id="sKey" title="검색설정" class="w120">
									<option value="">검색설정</option>
									<option value="NAME"<%=Fn_SetDefault(sKey,"NAME"," selected","")%>>이름</option>
									<!-- option value="ADDR"<%=Fn_SetDefault(sKey,"ADDR"," selected","")%>>거주지</option //-->
									<option value="MOBILE"<%=Fn_SetDefault(sKey,"MOBILE"," selected","")%>>핸드폰</option>
									<option value="PROJECT"<%=Fn_SetDefault(sKey,"PROJECT"," selected","")%>>최근지원한 프로젝트</option>
									<option value="CAREER"<%=Fn_SetDefault(sKey,"CAREER"," selected","")%>>경력</option>
								</select>
								<input type="text" id="sWord" name="sWord" value="<%=sWord%>" title="검색어" class="w190" />
								<input type="checkbox" name="sSearchIn" id="sSearchIn" value="Y" />
								<label for="sSearchIn">결과 내 재검색</label>
								<textarea name="sSearchHistory" id="sSearchHistory" style="display:none;"><%=sSearchHistory%></textarea>
								
								<span class="btns btn_b"><button type="submit">검색</button></span>
							</div>
							
							<% If Trim(strSearchInfor) <> "" Then %>
							<!-- 검색 키워드 //-->
							<div class="search_history"><strong>키워드 검색 내역 : </strong><%=strSearchInfor%></div>
							<% End If %>
						</div>

					</div>
					</form>
					
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="javascript:;" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnDel">회원삭제</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnBlack">블랙리스트</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnSearch">서칭거부</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnSMS">SMS발송</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnPrint">인쇄</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnXls">엑셀다운로드</a>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnReg">회원/이력서 수동입력</a>
							</span><span class="btns btn_g"><a href="javascript:alert('준비중입니다.');">ERP연동</a></span>
							</span><span class="btns btn_g"><a href="javascript:;" id="btnAgree">회원동의</a></span>
						</div>
					</div>
					
					<form name="frmList" id="frmList">
					<input type="hidden" id="Flag" name="Flag">
					<input type="hidden" id="SortXlsKey" name="SortXlsKey" value="<%=SortKey%>">
					<input type="hidden" id="SortXlsDir" name="SortXlsDir" value="<%=SortDir%>">
					<div class="tblArea">
						<table>
							<colgroup>
								<col class="w30" />
								<col class="w45" />
								<col class="w45" />
								<col class="w35" />
								<col class="w35" />
								<col />
								<col class="w100" />
								<col class="w100" />
								<col class="w100" />
								<col class="w30" />
								<col class="w120" />
								<col class="w85" />
								<col class="w85" />
								<col class="w200" />
								<col class="w100" />
								<col class="w45" />
							</colgroup>
							<thead>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th>선택</th>
									<th>번호</th>
									<th>이름</th>
									<th>
										<% If SortKey = "GENDER" And SortDir = "ASC" Then %>
										성별<a href="javascript:goSort('GENDER','DESC')">▲</a>
										<% Else %>
										성별<a href="javascript:goSort('GENDER','ASC')">▼</a>
										<% End If %>
									</th>
									<th>
										<% If SortKey = "AGE" And SortDir = "ASC" Then %>
										나이<a href="javascript:goSort('AGE','DESC')">▲</a>
										<% Else %>
										나이<a href="javascript:goSort('AGE','ASC')">▼</a>
										<% End If %>
									</th>
									<th>거주지</th>
									<th>휴대폰</th>
									<th>희망분야</th>
									<th>진행단계</th>
									<th>메모</th>
									<th>지원경로</th>
									<th>등록일</th>
									<th>수정일</th>
									<th>최근지원 프로젝트</th>
									<th>경력</th>
									<th>동의</th>
								</tr>
							</thead>
							<tbody>
								<% If Not kData Then %>
								<tr>
									<td colspan="15">등록된 회원이 없습니다.</td>
								</tr>
								<% Else %>
								<% 		
										Dim class_txt,CntStat 
										' TotalCnt(0), UsrSeq(1), UsrID(2), UsrName(3), Gender(4), Birthday(5), AddrSido(6)
										' AddrGugun(7), Mobile(8), HopePart(9), RegDate(10), ModDate(11), IsBlack(12), IsSearch(23)
										' CntMemo(13), Status(14), PartnerDepth1(15), PartnerDepth2(16), PrjName(17), EzCareer(18), CntEzCareer(19), HopePart2(20), RecrPart(21), RecrPart2(22), UsrType(23), IsSearch(24), CntStat(25)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
										
										
										if isNull(arrData(25,i))  then
										   CntStat = 0 
										else 
										   CntStat = arrData(25,i)
										end if 


										If CStr(Trim(arrData(12,i))) = "True" Then 
											class_txt = "blackList"
										ElseIf CStr(Trim(arrData(24,i))) = "True" Then 
											class_txt = "searchList"
										Else

											If CInt(CntStat ) > 0 then
												class_txt = "statList"
											Else
												class_txt = ""
											End If
										End If 

										
								%>
								<tr class="<%=class_txt%>" >
									<td><input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" /></td>
									<td><%=vNum%></td>
									<td><a href="mem_form.asp?Seq=<%=arrData(1,i)%>&Filter=<%=strFilter%>"><%=arrData(3,i)%></a></td>
									<td>
										<%
											Select Case arrData(4,i) 
												Case "1" : Response.Write "남" 
												Case "2" : Response.Write "여" 
												Case Else : Response.Write "-" 
											End Select 
										%>
									</td>
									<td><% If IsNumeric(Left(arrData(5,i),4)) Then %><%=Year(Date) - Left(arrData(5,i),4) + 1%><% Else %> - <% End If %></td>
									<td class="alignL"><%=arrData(6,i)%>&nbsp;<%=arrData(7,i)%></td>
									<td><%=arrData(8,i)%></td>
									<td>
										<%=arrData(21,i)%>
										<% If arrData(22,i) <> "" AND Not IsNull(arrData(22,i)) Then 	Response.Write " <br/>(" & arrData(22,i) & ")"%>
									</td>
									<td><%=arrData(14,i)%></td>
									<td><a href="javascript:goUserMemo(<%=arrData(1,i)%>);"><%=arrData(13,i)%></a></td>
									<td class="alignL"><%=arrData(15,i)%> 
										<% If Not FN_IsBlank(arrData(16,i)) Then %><br>(<%=arrData(16,i)%>)<% End If %>
									</td>
									<td><%=FN_SetDateTimeFormat(arrData(10,i),"YYYY-MM-DD")%></td>
									<td><%=FN_SetDateTimeFormat(arrData(11,i),"YYYY-MM-DD")%></td>
									<td><%=arrData(17,i)%></td>
									<td>
										<%
										If Not IsNull(arrData(19,i)) Then
											Response.Write arrData(18,i)
											If arrData(19,i) > 1 Then
												Response.Write "외 " & (arrData(19,i) - 1) & "개"
											End If
										End If	
										%>
									</td>
									<td>
										<%
											Select Case arrData(23,i) 
												Case "0" : Response.Write "미동의" 
												Case "1" : Response.Write "동의" 
												Case Else : Response.Write "미동의" 
											End Select 
										%>
									</td>
								</tr>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</tbody>
						</table>
						
						<div class="paging">
							<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
						</div>
					</div>
					<div class="regArea">
						<input type="hidden" name="PartnerDepth1" id="PartnerDepth1">
						<%
						Dim PartnerSql : PartnerSql = "Execute uspPartner_Info @Flag = 'CDD1'" 
						response.write " ": Call SB_MakeHTMLSelect(objDBCon, PartnerSql, "selPartnerD1", "ROOT", "", "", "", " class=""w100""", False)
						%>
						<span id="divPartner"></span>
						<%
						Dim RecruitSql : RecruitSql = "Execute uspRecruit_Info @Flag = 'ASLST'" 
						response.write " ": Call SB_MakeHTMLSelect(objDBCon, RecruitSql, "RecrSeq", "프로젝트 (채용공고) 선택", "", "", "", " class=""w340""", False)
						%>
						<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_APPLY", "Status", "RecrStatus", "진행단계", "", "", "", " class=""w105"" ", False)%>

						<span class="btns btn_b"><a href="javascript:;" id="btnRecr">채용공고 등록</a></span>
					</div>
					</form>
					
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>