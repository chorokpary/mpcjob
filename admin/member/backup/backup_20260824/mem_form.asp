<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_form.asp / 회원정보 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2014-01-07 / 강미현 / 4M / 희망분야 등록항목 삭제
'   #003 : 2014-12-16 / 이길홍 / 4M / 주민등록 등록항목 삭제
'____________________________________________________________________________________
%>
<%


	Dim Seq, Flag, Page, Filter
	Dim UsrName, Birthday, Gender, Mobile, Email, HopePart, HopePart2, UsrType
	Dim Zipcode1, Zipcode2, AddrSido, AddrGugun, Addr, AddrDetail, Jumin1, Jumin2
	Dim ImgPath, Attach1_Path, Attach2_Path, Attach3_Path, Attach1_FName, Attach2_FName, Attach3_FName
	Dim RegDate, ModDate, Intro, EduLast, EduType
	Dim kFields, arrLen, i, yy, mm
	Dim arrEdu, arrCareer, arrFam, arrMemo, arrApply, arrEzCareer
	Dim arrCdEzCr
	Dim strEzCareer, strAppCStyle, strAppCMsg
	Dim SidoSql, GugunSql
	Dim arrData, arrDataNum, sKey, sWord

	Dim ChargeCode

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)


	
	Call Main()
	Call getAdmList()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		Filter = FN_Req("Filter","")
		
		Dim arrFilter : arrFilter = Split(Filter,"[^]")
		
		Filter = ""
		For i = 0 To UBound(arrFilter)
			Filter = Filter & LEFT(arrFilter(i), InStr(arrFilter(i),"="))  &  Server.URLEncode( MID(arrFilter(i), InStr(arrFilter(i),"=")+1) ) & "[^]"	
		Next
		
		Call getCdData
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
		End If
	
	End Sub
	
	Sub getCdData()
		' 경력코드 DB Get
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfCode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "SEL"
			.Parameters("@Tbl").Value 	= "TBL_USER_CAREER_EZ"
			.Parameters("@Field").Value 	= "Career"
			.Parameters("@UseFlag").Value 	= "USE"
			
			' ## 코드 정보
			Set Rs = .Execute


			If Not Rs.eof Then
				arrCdEzCr = Rs.getrows
			End if

			Rs.Close
			Set Rs = Nothing
		End With
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				UsrName 	= FN_InputVal(Rs("UsrName"))
				Birthday 	= Rs("Birthday")
				Gender 	= Rs("Gender")
				Mobile 	= FN_InputVal(Rs("Mobile"))
				Email 	= FN_InputVal(Rs("Email"))
				'HopePart 	= Rs("HopePart")
				'HopePart2 = Rs("HopePart2")
				UsrType 	= Rs("UsrType")
				Zipcode1 = Rs("Zipcode1")
				Zipcode2 = Rs("Zipcode2")
				AddrSido 	= FN_InputVal(Rs("AddrSido"))
				AddrGugun = FN_InputVal(Rs("AddrGugun"))
				Addr 	= FN_InputVal(Rs("Addr"))
				AddrDetail = FN_InputVal(Rs("AddrDetail"))
				Jumin1 	= FN_InputVal(Rs("Jumin1"))
				Jumin2 	= FN_InputVal(Rs("Jumin2"))

				ImgPath 	= Rs("ImgPath")
				Attach1_Path = Rs("Attach1_Path")
				Attach1_FName = Rs("Attach1_FName")
				Attach2_Path = Rs("Attach2_Path")
				Attach2_FName = Rs("Attach2_FName")
				Attach3_Path = Rs("Attach3_Path")
				Attach3_FName = Rs("Attach3_FName")
				
				RegDate = Rs("RegDate")
				ModDate = Rs("ModDate")
				
				Intro = FN_BrToNL(Rs("Intro"))

				EduLast = Rs("EduLast")
				EduType = Rs("EduType")
				
				If FN_IsBlank(Addr) Then
					Addr = AddrSido & " " & AddrGugun
				End If
			End If

			Set Rs = Rs.NextRecordSet()
				
			'arrDataNum 	= UBound(arrData, 2)
			If Not Rs.Eof Then
				' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
				kFields 	= Array("UsrSeq","EduBYear","EduBMonth","EduEYear","EduEMonth","SchoolName","Major","Status","Local","Sort")
				arrEdu 	= Rs.GetRows(,,kFields)
			End If
			
			Set Rs = Rs.NextRecordSet()
				
			If Not Rs.Eof Then
				' UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10), CompanyYear(11)
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort","CompanyYear")
				arrCareer 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
				kFields 	= Array("UsrSeq","Relation","FmName","Age","Education","Job","IsLiveWith","Sort")
				arrFam 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' UsrMSeq(0), AdmName(1), Memo(2), RegDate(3)
				kFields 	= Array("UsrMSeq","AdmName","Memo","RegDate")
				arrMemo 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' PrjName(0), Title(1), AppBDate(2), AppEDate(3), RegDate(4), AdmName(5), CdVal(6), RecrPart(7), RecrPart2(8), CancelDate(9), RecrSeq(10), CntQues(11), Times(12)
				kFields 	= Array("PrjName","Title","AppBDate","AppEDate","RegDate","AdmName","CdVal","RecrPart","RecrPart2","CancelDate", "RecrSeq","CntQues","Times")
				arrApply 	= Rs.GetRows(,,kFields)
				
				HopePart = arrApply(7,0)
				HopePart2 = arrApply(8,0)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' Career(0)
				kFields 	= Array("Career")
				arrEzCareer 	= Rs.GetRows(,,kFields)
				
				strEzCareer = ","
				For i = 0 To UBound(arrEzCareer,2)
					strEzCareer = strEzCareer & arrEzCareer(0,i) & ","
				Next
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub

	Sub getAdmList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspAdm_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= "1"
			.Parameters("@per_page").Value 	= "50"
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@sLevel").Value 	= "1"
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1

		Else
			' TotalCnt(0), AdmSeq(1), AdmID(2), AdmName(3), AdmLevel(4), AdmEmail(5), VisitDate(6), AdmMobile(7), AdmTel(8)
			kFields 	= Array("TotalCnt","AdmSeq","AdmID","AdmName","AdmLevel","AdmEmail","VisitDate","AdmMobile","AdmTel")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub


	' ########## DB 연동 SELECT BOX 출력
	Sub SB_MakeHTMLSelect_recruit(argObjDB, strSQL, argName, argDefaultOpt, argDefaultVal, argSelectVal, argNullMsg, argScript, argIsUrlEncode)
	
		Dim fnRs
		
		Dim OptionData, OptionCheck, OptionRows
		Dim f
		
		set fnRs = argObjDB.Execute(strSQL)
	
		If fnRs.eof Then
			OptionCheck = False
		Else
			OptionCheck = True
			OptionData = fnRs.getrows
			OptionRows = ubound(OptionData,2)
		End if
			
		fnRs.close
		set fnRs=Nothing
		
		argScript = FN_StrInj(argScript)
		
		If OptionCheck Then
			response.write "<select id=""" & argName & """ name=""" & argName & """ " & argScript & ">" & chr(13)
			
			If argDefaultOpt <> "" Then
				response.write "<option value='" & argDefaultVal & "'>" & argDefaultOpt & "</option>" & chr(13)        
			End If
			
			For f=0 to OptionRows
		  
				If OptionData(1,f) <> "1546" and OptionData(1,f) <> "1612" and OptionData(1,f) <> "1619" Then 

					response.write "<option value='" 
					If argIsUrlEncode and OptionData(1,f) <> "" Then
						response.write Server.URLEncode(OptionData(1,f))
					Else
						response.write OptionData(1,f)
					End If
					response.write "'"
					If Trim(OptionData(1,f)) = Trim(argSelectVal) Then
						response.write " selected"
					End If
					response.write ">"& OptionData(3,f) &"</option>" & chr(13)

				End If 
		
			Next
			
			response.write "</select>"        
		Else
			If argNullMsg = "" Then
				response.write "<select id=""" & argName & """ name=""" & argName & """ " & argScript & ">" & chr(13)
				
				If argDefaultOpt <> "" Then
					response.write "<option value='" & argDefaultVal & "'>" & argDefaultOpt & "</option>" & chr(13)        
				End If
				
				response.write "</select>"        
			ElseIf argNullMsg = "[HIDE]" Then
			Else
				response.write argNullMsg
			End If
		End If
	
	End Sub

	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>회원정보 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript">
		var HopePart2 = "<%=HopePart2%>";

		function getAddrGugun(){
			var url = "/common/dev/ajax/make_select_gugun.asp";
			var param = "SIDO=" + $(this).val() + "&TargetID=AddrGugun&Css=w220";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divGugun").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function getAddrGugun2(){
			var url = "/common/dev/ajax/make_select_gugun.asp";
			var param = "SIDO=" + $(this).val() + "&TargetID=selAddrGugun&Css=w220";
			
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
			var param = "Parent=" + $(this).val() + "&TargetID=selPartnerD2_1";
		
			$("#selPartnerD1_1").val($(this).val());
			//$("#selPartnerD1_2").val($(this).val());
			$("#PartnerDepth1").val($(this).find("option:selected").text());
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ 
					$("#divPartner_1").html(content);
					//$("#divPartner_2").html(replaceAll(content,"selPartnerD2_1","selPartnerD2_2"));
					$("#PartnerDepth2").val($("#selPartnerD2_1").find("option:selected").text());
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function getPartnerRecruit(){
			var url = "/common/dev/ajax/make_select_recruit.asp";
			var param = "ChargeCode=" + $(this).val() + "&TargetID=Recruit_1";

			$("#AdminList_1").val($(this).val());
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ 
					$("#Recruit_1").html(content);
					var fOption = $("#Recruit_1").find("option").eq(0).val();
					$("#RecrSeq").val(fOption);
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function InitgetPartner(){

			var url = "/common/dev/ajax/make_select_partner.asp";
			var param = "Parent=5&TargetID=selPartnerD2_1";
	
			$("#selPartnerD1_1").val('5');
			$("#PartnerDepth1").val('온라인광고');
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ 
					$("#divPartner_1").html(content);
					//$("#divPartner_2").html(replaceAll(content,"selPartnerD2_1","selPartnerD2_2"));
					$("#PartnerDepth2").val($("#selPartnerD2_1").find("option:selected").text());
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
		
		function getPartner_Sub(){
			$("#selPartnerD2_1").val($(this).val());
			$("#selPartnerD2_2").val($(this).val());
			$("#PartnerDepth2").val($(this).find("option:selected").text());
		
		}
		
		function setRecr(){
			$("#RecrSeq").val($(this).val());
			$("#Recruit_1").val($(this).val());
			//$("#Recruit_2").val($(this).val());
		}
		
		function doRecruit(){
			if (!$("#PartnerDepth1").val()){
				alert("채용 Root를 선택해주세요.");
				return;
			}
			if (!$("#RecrSeq").val()){
				alert("프로젝트(채용공고)를 선택해주세요.");
				return;
			}

			$("#frmRecr").attr({action:"mem_proc.asp", method:"post"}).submit();
		}

		// 희망분야 2Depth
		function getHopePart2(){
			var url = "/common/dev/ajax/make_select_cd_2depth.asp";
			var param = "Flag=RecrPart&Parent=" + $("#HopePart").val() + "&TargetID=HopePart2&Def=" + HopePart2 + "&Css=w135";
			HopePart2 = "";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divHopePart2").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
		function chkForm(){
			var mobile  = $("#Mobile").val();
			if (mobile)	mobile = replaceAll(mobile,"-","");

			var orgMobile = $("#OrgMobile").val();
			if (orgMobile) orgMobile = replaceAll(orgMobile,"-","");
		
			$("#frmInfo").find("#Flag").val("<%=Flag%>");

			if (!$("#UsrName").val()){
				alert("이름을 입력해주세요.");
				$("#UsrName").focus();
				return;
			} else if (!$("#Birthday").val()){
				alert("생년월일을 입력해주세요.");
				$("#Birthday").focus();
				return;
		    } else if(!IsDate2($("#Birthday").val())){
				alert("생년월일은 [년도(2자리)월(2자리)일(2자리)] 형식으로 입력해주세요.");
		      	$("#Birthday").focus();
				return;
			/*
		    } else if(!IsDate($("#Birthday").val())){
				alert("생년월일은 [년도(4자리)-월(2자리)-일(2자리)] 형식으로 입력해주세요.");
		      	$("#Birthday").focus();
				return;
			*/
			} else if (!$("#Mobile").val()){
				alert("핸드폰번호를 입력해주세요.");
				$("#Mobile").focus();
				return;
			} else if (!IsMobile($("#Mobile").val(),"-")){
				alert("핸드폰번호가 바르지 않습니다.");
				$("#Mobile").focus();
				return;
			} else if ($("#ChkMobile").val()=="N"){
				if (mobile==orgMobile){
					$("#ChkMobile").val("Y");
				} else {
					if (!doChkMobile("submit")) {
						alert("사용중인 핸드폰번호 입니다.");
						return;
					}
				}
			}
		
			$("#frmInfo").attr({action:"https://www.mpcjob.co.kr/admin/member/mem_proc.asp", method:"post", enctype:"multipart/form-data"}).submit();
		}
		
		function doChkMobile(flag) {
			var rtnval = null;

			if (flag!="submit") {
				if (!$("#Mobile").val()){
					alert("핸드폰번호를 입력해주세요.");
					$("#Mobile").focus();
					return;
				} else if (!IsMobile($("#Mobile").val(),"-")){
					alert("핸드폰번호가 바르지 않습니다.");
					$("#Mobile").focus();
					return;
				}
			}

			var url = "/common/dev/ajax/check_exists_userid.asp";
			var param = "Mobile=" + $("#Mobile").val() ;

			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, async : false
				, success: function(result){ 
					if (result=="true"){
						$("#ChkMobile").val("Y");
						if (flag == "submit") {
							rtnval = true;
						}else{
							alert("사용 가능한 핸드폰번호 입니다.");
						}
					}else{
						$("#ChkMobile").val("N");
						if (flag == "submit") {
							rtnval = false;
						}else{
							alert("이미 사용중인 핸드폰번호 입니다.");
						}
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
			
			if (flag == "submit") {
				return rtnval;
			}
		}

		function doChangePass() {
			if (!$("#Passwd").val()){
				alert("비밀번호를 입력해주세요.");
				$("#Passwd").focus();
				return;
			}

			if(!fn_pw_check()){
				return false;
			}

			
			var url = "mem_proc.asp";
			var param = "IsAjax=1&Flag=PWD_C&Seq=<%=Seq%>&Passwd=" + $("#Passwd").val() ;
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(result){ 
					if (result=="SUCC"){
						alert("비밀번호가 변경되었습니다.");
						$("#Passwd").val("");
					}else{
						alert("처리중 오류가 발생했습니다.");
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
			
		}

		function doResetPass() {

			var url = "mem_proc.asp";
			var param = "IsAjax=1&Flag=PWD_R&Seq=<%=Seq%>" ;

			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(result){ 
					if (result=="SUCC"){
						alert("초기화 되었습니다.");
					}else{
						alert("처리중 오류가 발생했습니다.");
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
			
		}

		
		function doSendPass() {
			var url = "mem_proc.asp";
			var param = "IsAjax=1&Flag=PWD_S&Seq=<%=Seq%>&Mobile=" + $("#OrgMobile").val() ;
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(result){ 
					if (result=="SUCC"){
						alert("임시비밀번호가 회원의 핸드폰으로 전송되었습니다.");
						$("#Passwd").val("");
					}else{
						alert("처리중 오류가 발생했습니다.");
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
	
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}

		function goUserQuestion(seq){
			openWin("mem_question.asp?RecrSeq="+seq+"&Seq=<%=Seq%>","popQues",0,0,930,650,1);
		}
		
		function addDetailInfo(Flag){
			if (Flag=="Edu") {
				if (!$("#SchoolName").val()){
					alert("학교명을 입력해주세요.");
					$("#SchoolName").focus();
					return;
				}
			} else if (Flag=="Career"){
				if (!$("#CompanyName").val()){
					alert("회사명을 입력해주세요.");
					$("#CompanyName").focus();
					return;
				}
			} else if (Flag=="Fam"){
				if (!$("#Relation").val()){
					alert("가족 관계를 입력해주세요.");
					$("#Relation").focus();
					return;
				} else if (!$("#FmName").val()){
					alert("가족 성명을 입력해주세요.");
					$("#FmName").focus();
					return;
				} else if (!IsNumerics($("#Age").val())){
					alert("나이는 숫자만 입력할 수 있습니다.");
					$("#Age").focus();
					return;
				} else if (!$("#IsLiveWith").val()){
					alert("동거유무를 선택해주세요.");
					$("#IsLiveWith").focus();
					return;
				}
			} else if (Flag=="Memo"){
				if (!$("#Memo").val()){
					alert("메모를 입력해주세요.");
					$("#Memo").focus();
					return;
				}
			}

			$("#frmInfo").find("#Flag").val(Flag.toUpperCase());
			$("#frmInfo").find("#FlagSub").val("ADD");

			var url = "mem_proc.asp";
			var param = $("#frmInfo").serialize();
			param += "&IsAjax=1" ;
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(result){ 
					if (result=="SUCC"){
						getDetailInfo(Flag);
						if (Flag == "Memo")		$("#Memo").val("");
					}else{
						alert("처리중 오류가 발생했습니다.");
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});

		}
		
		function getDetailInfo(Flag){
			var url = "mem_info_ajax.asp";
			var param = "Flag=" + Flag.toUpperCase() + "&Seq=<%=Seq%>" ;

			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(contents){  $("#div" + Flag + "Info").html(contents); }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function delDetailInfo(Flag, DelSeq){
			if (DelSeq==0) {
				if (cntChecked("#div" + Flag + "Info")==0){
					alert("삭제할 데이터를 선택해 주세요.");
					return;
				}else{
					if(confirm("선택한 데이터를 삭제하시겠습니까?")){
						$("#frmInfo").find("#Flag").val(Flag.toUpperCase());
						$("#frmInfo").find("#FlagSub").val("DEL");
						DelSeq = getCheckedVal(Flag + "Sort");
					}else{
						return;
					}
				}
			}

			var url = "mem_proc.asp";
			var param = "IsAjax=1&Flag=" + Flag.toUpperCase() + "&FlagSub=DEL&Seq=<%=Seq%>&Sort=" + DelSeq;
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(result){ 
					if (result=="SUCC"){
						getDetailInfo(Flag);
					}else{
						alert("처리중 오류가 발생했습니다.");
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}
		
		function doShowHide(target){
			if ($("#" + target).is(':hidden')){
				$("#" + target).show();
			}else{
				$("#" + target).hide();
			}
		}
		
		function initPage(){
			// div hide 처리
			getHopePart2();
			InitgetPartner();
		}
		
		$(function() {

			$("#AddrSido").change(getAddrGugun);
			$("#selAddrSido").change(getAddrGugun2);
			$("#Recruit_1").change(setRecr);
			//$("#Recruit_2").change(setRecr);
			$("#selPartnerD1_1").change(getPartner);
			//$("#selPartnerD1_2").change(getPartner);
			$("#selPartnerD2_1").live("change",getPartner_Sub);
			//$("#selPartnerD2_2").live("change",getPartner_Sub);
			$("#HopePart").change(getHopePart2);
			$("#AdminList_1").change(getPartnerRecruit);
		    
			//$("#btnRecr_1").click(doRecruit);
			//$("#btnRecr_2").click(doRecruit);
			$("#btnChkMobile").click(doChkMobile);
			$("#btnFindZip").click(function(){ openWin("../popup/zipcode.asp?tgSi=AddrSido&tgGu=AddrGugun&tgZip1=Zipcode1&tgZip2=Zipcode2&tgAddr=Addr&tgAddr2=AddrDetail","popZipcode",0,0,550,300,0); });
			$("#btnPwdMod").click(doChangePass);
			$("#btnPwdSend").click(doSendPass);
			$("#btnPwdReSet").click(doResetPass);
			
			
			$("#btnPrint").click(function(){ openWin("mem_print.asp?Seq="+$("#Seq").val(),"popPrint",0,0,930,650,1) })
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); $("#pAttach1").show(); $("#pAttach2").show(); })
			$("#btnSubmit").click(chkForm);
			$("#btnRecr_1").click(chkForm);

			// 우편번호 대신 구군 시도로 적용
			$("#chkUseSido").click(function(){ if($(this).attr("checked")=="checked"){$("#divSido").show();}else{$("#divSido").hide();} })
			$("#btnPutSido").click(function(){
				$("#AddrSido").val($("#selAddrSido").val());
				$("#AddrGugun").val($("#selAddrGugun").val());
				$("#Zipcode1").val("");
				$("#Zipcode2").val("");
				$("#Addr").val($("#selAddrSido").val() + " " + $("#selAddrGugun").val());
			})
			
			

			$("#Mobile").keypress(function(e){ return OnlyNumber_H(e,false); });
			$("#Mobile").keyup(function(){ $(this).val( FillNumerics_H($(this).val())); $("#ChkMobile").val("N"); });

			$("#Jumin1").keypress(function(e){ return OnlyNumber(e,false); });
			$("#Jumin1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });

			$("#Jumin2").keypress(function(e){ return OnlyNumber(e,false); });
			$("#Jumin2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
			


			$("#sBDate").datepicker({
				date: $("#BDate").val()
				, current: $("#BDate").val()
			});
	
			$("#sEDate").datepicker({
				date: $("#EDate").val()
				, current: $("#EDate").val()
			});

		})
		

		$(window).load(initPage);
	</script>

	<script type="text/javascript">

     var pw_passed = true;  // 추후 벨리데이션 처리시에 해당 인자값 확인을 위해

     function pwInputValueCheck(){  //특정문자 체크
    	  var InputValue = ["password", "admin", " ", ",", ";", ":", "'", "\\" , "\"" , "|"];
    	  var validpw = $("#Passwd").val(); 
    	  for(i=0; i<InputValue.length; i++){
    	   if(validpw.indexOf(InputValue[i]) >= 0){
    	    return InputValue[i];
    	   }
    	  }
    	  return null;
    }


    function fn_pw_check() {
        
        var pw = $("#Passwd").val();

        pw_passed = true;

        if(pw.search(/₩s/) != -1){
        	  alert("비밀번호는 공백업이 입력해주세요.");
        	  return false;
         }

        /*
         * 비밀번호 체크
         * 영문 (52개) 특수문자(32개) 중 3가지 이상을 조합하여 2가지 이상일 경우 10글자 이상
         * 3가지 이상일 경우 8글자 이상
         */
         
        var pattern1 = /[0-9]/;
        var pattern3 = /[a-z]|[A-Z]/;
		var pattern4 = /[!@#$%^&*()<>]/;	// 특수문자


        var char_type = 0;
        
        if(pattern1.test(pw))  char_type = char_type+1;
        if(pattern3.test(pw))  char_type = char_type+1;
        if(pattern4.test(pw))  char_type = char_type+1;


       if((char_type <= 2)){
           alert("영문, 숫자, 특수문자를 조합해 8자리 이상 사용해 주세요.");
           return false;
       }else if(char_type > 2 && pw.length < 8 ){
    	   alert("영문, 숫자, 특수문자를 조합해 8자리 이상 사용해 주세요.");
           return false;
       }


        var SamePass_0 = 0; //동일문자 카운트
        var SamePass_1 = 0; //연속성(+) 카운드
        var SamePass_2 = 0; //연속성(-) 카운드

        for(var i=0; i < pw.length; i++) {
             var chr_pass_0;
             var chr_pass_1;
             var chr_pass_2;
             var chr_pass_3;
     
             if(i >= 2) {
            	 chr_pass_0 = pw.charCodeAt(i-3);
                 chr_pass_1 = pw.charCodeAt(i-2);
                 chr_pass_2 = pw.charCodeAt(i-1);
                 chr_pass_3 = pw.charCodeAt(i);
                 
                  //동일문자 카운트
                 if((chr_pass_0 == chr_pass_1) && (chr_pass_1 == chr_pass_2) && (chr_pass_2 == chr_pass_3)) {
                    SamePass_0++;
                  } 
                  else {
                   SamePass_0 = 0;
                   }

                  //연속성(+) 카운드
                 if((chr_pass_0 - chr_pass_1 == 1) && (chr_pass_1 - chr_pass_2 == 1)  && (chr_pass_2 - chr_pass_3 == 1)) {
                     SamePass_1++;
                  }
                  else {
                   SamePass_1 = 0;
                  }
          
                  //연속성(-) 카운드
                 if((chr_pass_0 - chr_pass_1 == -1) && (chr_pass_1 - chr_pass_2 == -1) && (chr_pass_2 - chr_pass_3 == -1)) {
                     SamePass_2++;
                  } 
                  else {
                   SamePass_2 = 0;
                  }  
             }     
              
            if(SamePass_0 > 0) {
               alert("동일 또는 연속된 문자/숫자 4자리 이상은 사용하실 수 없습니다.");
               pw_passed=false;
             }

            if(SamePass_1 > 0 || SamePass_2 > 0 ) {
               alert("동일 또는 연속된 문자/숫자 4자리 이상은 사용하실 수 없습니다.");
               pw_passed=false;
             } 
             
             if(!pw_passed) {             
                  return false;
            }
        }

        var srt = pwInputValueCheck();
        if(srt != null){
       	   alert("특정 특수문자 및 특정단어는 비밀번호로 사용하실 수 없습니다.");
       	   return false;
       	  }

        return true;
    }
</script>

</head>

<body>
	<div id="wrap" class="memInfo">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>

				<div id="contArea" class="<%=Fn_SetDefault(Flag,"ADD","job_reg","mem_info")%>">
					<h2>회원정보</h2>
					<% If Flag = "MOD" Then %>
					<div class="btnArea mt-20">
						<div class="fl_r">
							<span class="btns btn_g"><a href="#btnPrint" id="btnPrint">인쇄하기</a></span>
						</div>
					</div>
					<% End If %>

					<form id="frmInfo" name="frmInfo" action="mem_proc.asp" method="post" enctype!="multipart/form-data" >
					<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
					<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
					<input type="hidden" id="Filter" name="Filter" value="<%=FN_InputVal(Filter)%>">
					<input type="hidden" id="FlagSub" name="FlagSub">
					<input type="hidden" name="txtID" id="txtID"  value="<%=Replace(Mobile,"-","")%>"/>
					<!--
					<input type="hidden" id="AddPartnerD1" name="AddPartnerD1">
					<input type="hidden" id="AddPartnerD2" name="AddPartnerD2">
					<input type="hidden" id="AddRecrSeq" name="AddRecrSeq">
					//-->
					<!-- 채용공고 등록용 //-->
					<!--
					<input type="hidden" id="Flag" name="Flag" value="RECR">
					 //-->
					<input type="hidden" id="PartnerDepth1" name="PartnerDepth1">
					<input type="hidden" id="PartnerDepth2" name="PartnerDepth2">
					<input type="hidden" id="RecrSeq" name="RecrSeq">
					
					<div class="searchArea">

						<%
						'Dim AdmListSql : AdmListSql = "Execute uspAdm_List @Flag = 'ADM', @pno = '1', @per_page = '50', @sKey = '', @sWord = '', @sLevel = '1' " 
						'Dim AdmSql : AdmSql = "Execute uspAdm_Info @Flag = 'LV1'" 
						'response.write " ": Call SB_MakeHTMLSelect_recruit(objDBCon, AdmSql, "AdminList_1", "담당자", "", Session("ASeq"), "", " class=""w100""", False)

						Dim AdmSql : AdmSql = "Execute uspAdm_Info @Flag = 'LV1'" 
						Response.Write " ": Call SB_MakeHTMLSelect(objDBCon, AdmSql, "AdminList_1", "담당자", "", Session("ASeq"), "", " class=""w100""", False)
						%>

						<%
						Dim PartnerSql : PartnerSql = "Execute uspPartner_Info @Flag = 'CDD1'" 
						response.write " ": Call SB_MakeHTMLSelect(objDBCon, PartnerSql, "selPartnerD1_1", "ROOT", "", "", "", " class=""w100""", False)
						%>
						<span id="divPartner_1"></span>
						<%
						Dim RecruitSql : RecruitSql = "Execute uspRecruit_Info2 @Flag = 'ASLST2', @ChargeCode ='"&Session("ASeq")&"' " 
						response.write " ": Call SB_MakeHTMLSelect(objDBCon, RecruitSql, "Recruit_1", "프로젝트 (채용공고) 선택", "", "", "", " class=""w340""", False)
						%>
						<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_APPLY", "Status", "RecrStatus", "진행단계", "", "", "", " class=""w105"" ", False)%>
						
						<% If Flag = "MOD" Then %>
						<span class="btns btn_b"><a href="javascript:;" id="btnRecr_1">등록하기</a></span>
						<% End If %>
					</div>

					<% If Flag = "ADD" Then %>
					<div class="btnArea">
						<div class="fl_l">
							<p><span class="ess">*</span> 항목은 필수 항목입니다.</p>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#" id="btnSubmit">이력서 등록</a></span>
						</div>
					</div>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w85" />
								<col />
								<col class="w85" />
								<col class="w150" />
								<col class="w85" />
								<col class="w150" />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="UsrName"><span class="ess">*</span> 이름</label></th>
									<td>
										<input type="text" id="UsrName" name="UsrName" title="이름" value="" class="w200" maxlength="25" /> 
									</td>
									<th><label for="Birthday"><span class="ess">*</span> 생년월일</label></th>
									<td>
										<select name="Birthday_Y" id="Birthday_Y" title="년" class="w55">
											<option value="1900" selected>19</option>
											<option value="2000">20</option>
										</select>
										<input type="text" id="Birthday" name="Birthday" title="생년월일" value="" class="w65" maxlength="6" style="ime-mode:disabled;" />
									</td>
									<th><label for="Gender"><span class="ess">*</span> 성별</label></th>
									<td>
										<select id="Gender" name="Gender" title="성별선택" class="w65">
											<option value="2">여</option>
											<option value="1">남</option>
										</select>
									</td>
								</tr>
								<tr>
									<th><label for="Mobile"><span class="ess">*</span> 핸드폰번호</label></th>
									<td>
										<input type="tel" id="Mobile" name="Mobile" title="핸드폰번호" value="" class="w200" maxlength="13" style="ime-mode:disabled;" /> 
										<span class="btns btn_in"><a href="#btnChkMobile" id="btnChkMobile">중복체크</a></span>
										<input type="hidden" id="ChkMobile" name="ChkMobile" value="N">
									</td>
									<th><label for="HopePart">희망분야</label></th>
									<td colspan="3">
										<!--
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "RecrPart", "HopePart", "-- 희망분야 --", "", HopePart, "등록된 분야가 없습니다.", " class=""w135"" ", False)%>
										<span id="divHopePart2"></span>
										<%' response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER", "HopePart", "HopePart_OLD", "없음", "", "", "", " class=""w290"" ", False)%>
										//-->
										<%=HopePart%>
										<% If HopePart2 <> "" AND Not IsNull(HopePart2) Then 	Response.Write "(" & HopePart2 & ")"%>
									</td>
								</tr>
								<tr>
									<th><label for="AddrSido"><span class="ess"></span> 주소</label></th>
									<td colspan="5">
										<%
										SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY CASE SIDO WHEN '서울' THEN 0 ELSE 1 END ASC, SIDO ASC" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "AddrSido", "", "", "", "", " class=""w160"" ", False)
										%>
										<span id="divGugun"> 
										<%
										GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '서울' GROUP BY GUGUN ORDER BY GUGUN" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, GugunSql, "AddrGugun", "지역중분류", "", "", "", " class=""w220"" ", False)
										%>
										</span>
									</td>
								</tr>
								<tr>
									<th><label for="Attach_1">첨부파일 1</label></th>
									<td colspan="5">
										<input type="file" id="Attach_1" name="Attach_1" title="첨부파일 1" value="" class="w660" size="95" /> 
									</td>
								</tr>
								<tr>
									<th><label for="Attach_2">첨부파일 2</label></th>
									<td colspan="5">
										<input type="file" id="Attach_2" name="Attach_2" title="첨부파일 2" value="" class="w660" size="95" /> 
									</td>
								</tr>
								<tr>
									<th><label for="Attach_3">첨부파일 3</label></th>
									<td colspan="5">
										<input type="file" id="Attach_3" name="Attach_3" title="첨부파일 3" value="" class="w660" size="95" /> 
									</td>
								</tr>
								<% If  Ubound(arrCdEzCr,2) >= 0 Then  %>
								<tr>
									<th><label for="Career_1">경력</label></th>
									<td colspan="5">
										<% For i = 0 To Ubound(arrCdEzCr,2) %>
										<span class="careerList"><input type="checkbox" id="EzCareer_<%=i+1%>" name="EzCareer" value="<%=FN_InputVal(arrCdEzCr(1,i))%>"><label for="EzCareer_<%=i+1%>" class="ml5"><%=arrCdEzCr(1,i)%></label></span>
										<% Next %>
									</td>
								</tr>
								<% End If %>
							</tbody>
						</table>
					</div>
					<% ElseIf Flag = "MOD" Then %>
					<div class="section">
						<div class="tblArea">
							<table class="tblForm">
								<colgroup>
									<col class="w85" />
									<col />
									<col class="w85" />
									<col class="w150" />
									<col class="w85" />
									<col class="w150" />
								</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td colspan="3"></td>
									</tr>
									<tr>
										<th><label for="">등록일</label></th>
										<td><%=RegDate%></td>
										<th><label for="">수정일</label></th>
										<td colspan="3"><%=ModDate%></td>
									</tr>
									<tr>
										<th><label for="UsrName"><span class="ess">*</span> 이름</label></th>
										<td>
											<input type="text" id="UsrName" name="UsrName" title="이름" value="<%=UsrName%>" class="w200" maxlength="25" /> 
										</td>
										<th><label for="Birthday">생년월일</label></th>
										<td>
											<select name="Birthday_Y" id="Birthday_Y" title="년" class="w55">
												<option value="1900" <%If Left(Birthday,2) = "19" Or Left(Birthday,2) = "" Then %>selected<%End If%>>19</option>
												<option value="2000" <%If Left(Birthday,2) = "20" Then %>selected<%End If%>>20</option>
											</select>

											<input type="text" id="Birthday" name="Birthday" title="생년월일" value="<%=Right(Replace(Birthday,"-",""),6)%>" class="w65" maxlength="6" style="ime-mode:disabled;" />
											<!-- input type="text" id="Birthday" name="Birthday" title="생년월일" value="<%=Birthday%>" class="w220" maxlength="10" style="ime-mode:disabled;" //-->
											<!--select id="" class="w60">
												<option value=""></option>
											</select>
											년
											<select id="" class="w50">
												<option value=""></option>
											</select>
											월
											<select id="" class="w50">
												<option value=""></option>
											</select>
											일 //-->
										</td>
										<th><label for="Gender">성별</label></th>
										<td>
											<select id="Gender" name="Gender" title="성별선택" class="w65">
												<option value="1"<%=Fn_SetDefault(Gender,"1"," selected","")%>>남</option>
												<option value="2"<%=Fn_SetDefault(Gender,"2"," selected","")%>>여</option>
											</select>
										</td>
									</tr>
									<tr>
										<th><label for=""><span class="ess">*</span> 아이디</label></th>
										<td><%=oSeed.Decrypt(Mobile)%></td>
										<th><label for="HopePart">희망분야</label></th>
										<td colspan="3">
											<!--
											<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "RecrPart", "HopePart", "-- 희망분야 --", "", HopePart, "등록된 분야가 없습니다.", " class=""w135"" ", False)%>
											<span id="divHopePart2"></span>
											<%' response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER", "HopePart", "HopePart", "없음", "", HopePart, "", " class=""w290"" ", False)%>
											//-->
											<%=HopePart%>
											<% If HopePart2 <> "" AND Not IsNull(HopePart2) Then 	Response.Write "(" & HopePart2 & ")"%>
										</td>
									</tr>
									<!-- <tr>
										<th><label for="Jumin1">주민번호</label></th>
										<td colspan="5">
											<input type="number" id="Jumin1" name="Jumin1" title="주민등록번호 앞 6자리" value="<%=Jumin1%>" maxlength="6" class="w80" style="ime-mode:disabled;" />
											-
											<input type="number" id="Jumin2" name="Jumin2" title="주민등록번호 뒤 7자리" value="<%=Jumin2%>" maxlength="7" class="w80" style="ime-mode:disabled;" />
										</td>
									</tr> -->
									<tr>
										<th><label for="pImgPath">이미지</label></th>
										<td colspan="5">
											<% If Not FN_isBlank(ImgPath) Then %>
											<p id="pImgPath" class="filename"><img src="<%=ImgPath%>" width="116" height="144"> <span class="btns btn_in"><a href="javascript:doDelFile('pImgPath','DelImgPath','<%=ImgPath%>')">삭제</a></span></p>
											<% Else %>
											<p class="filename"><span>등록된 사진이 없습니다.</span></p>
											<% End If %>
											<input type="file" id="ImgPath" name="ImgPath" title="이미지첨부" value="" class="w660" size="95" /> 
											<input type="hidden" id="DelImgPath" name="DelImgPath" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Attach_1">첨부파일 1</label></th>
										<td colspan="5">
											<% If Not FN_isBlank(Attach1_Path) Then %>
											<p id="pAttach_1"><a href="javascript:goDown('<%=Seq%>','USERS','Attach1_Path','<%=Attach1_FName%>');"><%=Attach1_FName%></a> <span class="btns btn_in"><a href="javascript:doDelFile('pAttach_1','DelAttach_1','<%=Attach1_Path%>')">삭제</a></span></p>
											<% End If %>
											<input type="file" id="Attach_1" name="Attach_1" title="첨부파일 1" value="" class="w660" size="95" /> 
											<input type="hidden" id="DelAttach_1" name="DelAttach_1" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Attach_2">첨부파일 2</label></th>
										<td colspan="5">
											<% If Not FN_isBlank(Attach2_Path) Then %>
											<p id="pAttach_2"><a href="javascript:goDown('<%=Seq%>','USERS','Attach2_Path','<%=Attach2_FName%>');"><%=Attach2_FName%></a> <span class="btns btn_in"><a href="javascript:doDelFile('pAttach_2','DelAttach_2','<%=Attach2_Path%>')">삭제</a></span></p>
											<% End If %>
											<input type="file" id="Attach_2" name="Attach_2" title="첨부파일 2" value="" class="w660" size="95" /> 
											<input type="hidden" id="DelAttach_2" name="DelAttach_2" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Attach_3">첨부파일 3</label></th>
										<td colspan="5">
											<% If Not FN_isBlank(Attach3_Path) Then %>
											<p id="pAttach_3"><a href="javascript:goDown('<%=Seq%>','USERS','Attach3_Path','<%=Attach3_FName%>');"><%=Attach3_FName%></a> <span class="btns btn_in"><a href="javascript:doDelFile('pAttach_3','DelAttach_3','<%=Attach3_Path%>')">삭제</a></span></p>
											<% End If %>
											<input type="file" id="Attach_3" name="Attach_3" title="첨부파일 3" value="" class="w660" size="95" /> 
											<input type="hidden" id="DelAttach_3" name="DelAttach_3" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Passwd"><% If UsrType = "0" Then %><span class="ess">*</span><% End If %> 비밀번호</label></th>
										<td>
											<input type="password" id="Passwd" name="Passwd" title="비밀번호" value="" class="w80" maxlength="30" /> 
											<span class="btns btn_in"><a href="#btnPwdMod" id="btnPwdMod">수정</a>
											</span><span class="btns btn_in"><a href="#btnPwdSend" id="btnPwdSend">임시비밀번호 전송</a></span><span class="btns btn_in"><a href="#btnPwdReSet" id="btnPwdReSet">초기화</a></span>
										</td>
										<th><label for="Email"><span class="ess">*</span> E-mail 주소</label></th>
										<td colspan="3">
											<input type="email" id="Email" name="Email" title="E-mail 주소" value="<%=Email%>" class="w285" maxlength="100" style="ime-mode:disabled;" /> 
										</td>
									</tr>
									<tr>
										<th><label for="Mobile"><span class="ess">*</span> 핸드폰번호</label></th>
										<td colspan="5">
											<input type="tel" id="Mobile" name="Mobile" title="핸드폰번호" value="<%If Mobile <> "" And Flag = "MOD" Then Response.write oSeed.Decrypt(Mobile) End If %>" class="w285" maxlength="13" style="ime-mode:disabled;" /> 
											<span class="btns btn_in"><a href="#btnChkMobile" id="btnChkMobile">중복체크</a></span>
											<input type="hidden" id="ChkMobile" name="ChkMobile" value="Y">
											<input type="hidden" id="OrgMobile" name="OrgMobile" value="<%If Mobile <> "" And Flag = "MOD" Then Response.write oSeed.Decrypt(Mobile) End If %>">
										</td>
									</tr>
									<tr>
										<th><label for="Zipcode1">우편번호</label></th>
										<td colspan="5">
											<input type="number" id="Zipcode1" name="Zipcode1" title="우편번호" value="<%=Zipcode1%>" maxlength="3" class="w80" readonly />
											-
											<input type="number" id="Zipcode2" name="Zipcode2" title="우편번호" value="<%=Zipcode2%>" maxlength="3" class="w80" readonly />
											<span class="btns btn_in"><a href="javascript:;" id="btnFindZip">우편번호 찾기</a></span>
											<input type="checkbox" id="chkUseSido"><label for="chkUseSido">시도로 검색</label>
											
											<div id="divSido" style="display:none;">
												<%
												SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY CASE SIDO WHEN '서울' THEN 0 ELSE 1 END ASC, SIDO ASC" 
												response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "selAddrSido", "", "", "", "", " class=""w160"" ", False)
												%>
												<span id="divGugun"> 
												<%
												 GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '서울' GROUP BY GUGUN ORDER BY GUGUN" 
												response.write " ": Call SB_MakeHTMLSelect(objDBCon, GugunSql, "selAddrGugun", "지역중분류", "", "", "" , " class=""w220"" ", False)
												%>
												</span>
												<span class="btns btn_in"><a href="javascript:;" id="btnPutSido">적용</a></span>
											</div>
										
											<input type="hidden" id="AddrSido" name="AddrSido" value="<%=AddrSido%>">
											<input type="hidden" id="AddrGugun" name="AddrGugun" value="<%=AddrGugun%>">
										</td>
									</tr>
									<tr>
										<th><label for="Addr">주소</label></th>
										<td colspan="5">
											<input type="text" id="Addr" name="Addr" title="주소" value="<%=Addr%>" class="w675" readonly maxlength="75" />
										</td>
									</tr>
									<tr>
										<th><label for="AddrDetail">상세주소</label></th>
										<td colspan="5">
											<input type="text" id="AddrDetail" name="AddrDetail" title="상세주소" value="<%=AddrDetail%>" class="w675" maxlength="75" />
										</td>
									</tr>
									<% If  Ubound(arrCdEzCr,2) >= 0 Then  %>
									<tr>
										<th><label for="Career_1">경력</label></th>
										<td colspan="5">
											<% For i = 0 To Ubound(arrCdEzCr,2) %>
											<span class="careerList"><input type="checkbox" id="EzCareer_<%=i+1%>" name="EzCareer" value="<%=FN_InputVal(arrCdEzCr(1,i))%>" <%=Fn_SetDefault(InStr(strEzCareer,"," & arrCdEzCr(1,i) & ","), 0, "", " checked")%>><label for="EzCareer_<%=i+1%>" class="ml5"><%=arrCdEzCr(1,i)%></label></span>
											<% Next %>
										</td>
									</tr>
									<% End If %>
								</tbody>
							</table>
						</div>
					</div>
					<div class="section">
						<h2>학력사항</h2>
						<a href="javascript:doShowHide('divEduArea')" class="more">더보기</a>
					
						<div id="divEduArea" class="tblArea" style="display:none;">
							<div id="divEduInfo">
								<%
								Dim strFinEdu : strFinEdu = ""
								If IsArray(arrEdu) Then
									arrLen = UBound(arrEdu, 2)
									strFinEdu = arrEdu(5,arrLen)
								Else
									arrLen = -1
								End If
								%>
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<!-- <col class="w30" />-->
										<col class="w110" /> 
										<col class="w110" />
										<col />
										<!-- <col class="w100" />
										<col class="w80" />
										<col class="w100" /> -->
									</colgroup>
									<tbody>
										<tr class="design">
											<td></td>
											<td></td>
											<td></td>
											<!-- <td></td>
											<td></td>
											<td></td>
											<td></td>
											<td></td> -->
										</tr>
										<tr>
											<th rowspan="<%=(4+arrLen)%>">최종학력사항</th>
											<td class="alignC">
												<select id="EduLast" name="EduLast" title="최종학력" class="w120">
													<option value="0" <%=Fn_SetDefault(EduLast,"0"," selected","")%>>중학교</option>
													<option value="1" <%=Fn_SetDefault(EduLast,"1"," selected","")%>>고등학교</option>
													<option value="2" <%=Fn_SetDefault(EduLast,"2"," selected","")%>>대학(2,3년)</option>
													<option value="3" <%=Fn_SetDefault(EduLast,"3"," selected","")%>>대학교(4년)</option>
													<option value="4" <%=Fn_SetDefault(EduLast,"4"," selected","")%>>대학원</option>
												</select>
											</td>
											<td class="alignC">
												&nbsp;<input type="radio" name="EduType" id="EduType1" value="1" <%=Fn_SetDefault(EduType,"1"," checked","")%> >졸업
												<input type="radio" name="EduType" id="EduType2" value="2" <%=Fn_SetDefault(EduType,"2"," checked","")%>>재학중
												<input type="radio" name="EduType" id="EduType3" value="3" <%=Fn_SetDefault(EduType,"3"," checked","")%>>휴학
												<input type="radio" name="EduType" id="EduType4" value="4" <%=Fn_SetDefault(EduType,"4"," checked","")%>>중퇴
											</td>
											<!-- <th colspan="7" class="alignL"><% If arrLen > -1 Then %>최종학력 : <%=strFinEdu%> <% End If %></th> -->
										</tr>
										<!-- <tr>
											<th></th>
											<th>입학년월</th>
											<th>졸업년월</th>
											<th>학교명</th>
											<th>전공</th>
											<th>졸업여부</th>
											<th>소재지</th>
										</tr> -->
										<% 		
										' UsrSeq(0), EduBYear(1), EduBMonth(2), EduEYear(3), EduEMonth(4), SchoolName(5), Major(6), Status(7), Local(8), Sort(9)
										For i = 0 To arrLen 
										%>
										<!-- <tr>
											<td class="alignC"><input type="checkbox" id="EduSort_<%=arrEdu(9,i)%>" name="EduSort" value="<%=arrEdu(9,i)%>" title="선택" /></td>
											<td class="alignC"><% If arrEdu(1,i) <> "" And arrEdu(2,i) <> "" Then %> <%=arrEdu(1,i)%> 년 <%=arrEdu(2,i)%> 월 <% End If %></td>
											<td class="alignC"><% If arrEdu(3,i) <> "" And arrEdu(4,i) <> "" Then %> <%=arrEdu(3,i)%> 년 <%=arrEdu(4,i)%> 월 <% End If %></td>
											<td><%=arrEdu(5,i)%></td>
											<td class="alignC"><%=arrEdu(6,i)%></td>
											<td class="alignC"><%=arrEdu(7,i)%></td>
											<td class="alignC"><%=arrEdu(8,i)%></td>
										</tr> -->
										<% Next %>
										<!-- <tr>
											<td class="alignC"></td>
											<td class="alignC">
												<select id="EduBYear" name="EduBYear" title="입학년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) + 3 %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="EduBMonth" name="EduBMonth" title="입학월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
											</td>
											<td class="alignC">
												<select id="EduEYear" name="EduEYear" title="졸업년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) + 3 %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="EduEMonth" name="EduEMonth" title="졸업월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
											</td>
											<td><input type="text" id="SchoolName" name="SchoolName" title="학교명" value="" class="w145" maxlength="20" /></td>
											<td class="alignC"><input type="text" id="Major" name="Major" title="전공" value="" class="w80" maxlength="10" /></td>
											<td class="alignC">
												<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Status", "Status", "", "", "", "", " class=""w55"" ", False)%>
											</td>
											<td class="alignC">
												<% response.write " ": Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_USER_EDUCATION", "Local", "Local", "", "", "", "", " class=""w55"" ", False)%>
											</td> -->
										</tr>
									</tbody>
								</table>
							</div>
							<!-- <div class="fl_r">
								<span class="btns btn_b"><a href="javascript:addDetailInfo('Edu')">추가</a></span><span class="btns btn_b"><a href="javascript:delDetailInfo('Edu',0)">삭제</a></span>
							</div> -->
						</div>
					</div>
					
					<div class="section">
						<h2>경력사항</h2>
						<a href="javascript:doShowHide('divCareerArea')" class="more">더보기</a>
						<div id="divCareerArea" class="tblArea" style="display:none;">
							<div id="divCareerInfo">
								<%
								If IsArray(arrCareer) Then
									arrLen = UBound(arrCareer, 2)
								Else
									arrLen = -1
								End If
								%>
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col class="w30" />
										<col class="w340" />
										<!-- <col class="w100" /> -->
										<col class="w100" />
										<col class="w100" />
										<col />
										<col class="w100" />
									</colgroup>
									<tbody>
										<tr class="design">
											<td></td>
											<td></td>
											<td></td>
											<td></td>
											<!-- <td></td> -->
											<td></td>
											<td></td>
											<td></td>
										</tr>
										<tr>
											<th rowspan="<%=(3+arrLen)%>">경력사항</th>
											<th></th>
											<th>근무기간</th>
											<th>직장명</th>
											<!-- <th>직위</th> -->
											<th>부서명</th>
											<th>담당업무</th>
											<th>퇴직사유</th>
										</tr>
										<% 		
										'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10)
										For i = 0 To arrLen 
										%>
										<tr>
											<td class="alignC"><input type="checkbox" id="CareerSort_<%=arrCareer(10,i)%>" name="CareerSort" value="<%=arrCareer(10,i)%>" title="선택" /></td>
											<td class="alignC">
												
												<input type="text"  title="입사년도" value="<% If arrCareer(1,i) <> "" And arrCareer(2,i) <> "" Then %><%=arrCareer(1,i)%><%=arrCareer(2,i)%><% End If %>" class="w50" readonly />~
												<input type="text"  title="입사년도" value="<% If arrCareer(3,i) <> "" And arrCareer(4,i) <> "" Then %><%=arrCareer(3,i)%><%=arrCareer(4,i)%><% End If %>" class="w50" readonly />
											<% 
											
											If arrCareer(1,i) <> "" And arrCareer(3,i) <> ""  Then 
												 Dim mon
												 mon = CInt(DateDiff("M", ""&arrCareer(1,i)&"-"&arrCareer(2,i)&"-01",""&arrCareer(3,i)&"-"&arrCareer(4,i)&"-01"))
											%>
												<input type="text" title="입사년도" value="<%=mon%>개월" class="w50" readonly />
											<% Else %>
												<input type="text"  title="입사년도" value="<%
												Select Case(arrCareer(11,i))
													Case "0"
														Response.Write "1년미만"
													Case "1"
														Response.Write "1년이상"
													Case "2"
														Response.Write "2년이상"
													Case "3"
														Response.Write "3년이상"
													Case "4"
														Response.Write "4년이상"
												End Select 
												%>" class="w50" readonly />
											<% End If %>
											</td>
											<td><%=arrCareer(5,i)%></td>
											<!-- <td class="alignC"><%=arrCareer(7,i)%></td> -->
											<td class="alignC"><%=arrCareer(6,i)%></td>
											<td><%=arrCareer(8,i)%></td>
											<td><%=arrCareer(9,i)%></td>
										</tr>
										<% Next %>
										<tr>
											<td class="alignC"></td>
											<td class="alignC">
												<select id="CrBYear" name="CrBYear" title="입사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrBMonth" name="CrBMonth" title="입사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
												~
												<select id="CrEYear" name="CrEYear" title="퇴사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrEMonth" name="CrEMonth" title="퇴사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>

												<!-- <select id="CompanyYear" name="CompanyYear" title="근무개월수" class="w70">
													<option value="0">1년미만</option>
													<option value="1">1년이상</option>
													<option value="2">2년이상</option>
													<option value="3">3년이상</option>
													<option value="4">4년이상</option>
												</select>  -->
											</td>
											<td><input type="text" id="CompanyName" name="CompanyName" title="직장명" value="" class="w80" maxlength="20" /></td>
											<!-- <td class="alignC"><input type="text" id="Position" name="Position" title="직위" value="" class="w50" maxlength="20" /></td> -->
											<td class="alignC"><input type="text" id="Department" name="Department" title="부서" value="" class="w60" maxlength="15" /></td>
											<td><input type="text" id="Task" name="Task" name="Task" title="업무" value="" class="w105" maxlength="15" /></td>
											<td><input type="text" id="Reason" name="Reason" name="Reason" title="퇴사사유" value="" class="w45" maxlength="25" /></td>
										</tr>
									</tbody>
								</table>
							</div>
							<div class="fl_r">
								<span class="btns btn_b"><a href="javascript:addDetailInfo('Career')">추가</a></span><span class="btns btn_b"><a href="javascript:delDetailInfo('Career',0)">삭제</a></span>
							</div>
						</div>
					</div>
					
					<!-- <div class="section">
						<h2>가족사항</h2>
						<a href="javascript:doShowHide('divFamArea')" class="more">더보기</a>
						<div id="divFamArea" class="tblArea" style="display:none;">
							<div id="divFamInfo">
								<%
								If IsArray(arrFam) Then
									arrLen = UBound(arrFam, 2)
								Else
									arrLen = -1
								End If
								%>
								<table class="tblForm">
									<colgroup>
										<col class="w85" />
										<col class="w30" />
										<col class="w80" />
										<col class="w80" />
										<col class="w80" />
										<col class="w80" />
										<col />
										<col class="w80" />
									</colgroup>
									<tbody>
										<tr class="design">
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
											<th rowspan="<%=(3+arrLen)%>">가족사항</th>
											<th></th>
											<th>관계</th>
											<th>성명</th>
											<th>연령</th>
											<th>학력</th>
											<th>근무처</th>
											<th>동거유무</th>
										</tr>
										<% 		
										' UsrSeq(0), Relation(1), FmName(2), Age(3), Education(4), Job(5), IsLiveWith(6), Sort(7)
										For i = 0 To arrLen 
										%>
										<tr>
											<td class="alignC"><input type="checkbox" id="FamSort_<%=arrFam(7,i)%>" name="FamSort" value="<%=arrFam(7,i)%>" title="선택" /></td>
											<td class="alignC"><%=arrFam(1,i)%></td>
											<td class="alignC"><%=arrFam(2,i)%></td>
											<td class="alignC"><%=arrFam(3,i)%></td>
											<td><%=arrFam(4,i)%></td>
											<td><%=arrFam(5,i)%></td>
											<td class="alignC"><%=Fn_SetDefault(arrFam(6,i),True,"유","무")%></td>
										</tr>
										<% Next %>
										<tr>
											<td class="alignC"></td>
											<td class="alignC"><input type="text" id="Relation" name="Relation" title="관계" value="" class="w60" maxlength="5" /></td>
											<td class="alignC"><input type="text" id="FmName" name="FmName" title="성명" value="" class="w60" maxlength="10" /></td>
											<td class="alignC"><input type="text" id="Age" name="Age" title="연령" value="" class="w60" maxlength="3" /></td>
											<td><input type="text" id="Education" name="Education" title="학력" value="" class="w60" maxlength="15" /></td>
											<td><input type="text" id="Job" name="Job" title="근무처" value="" class="w245" maxlength="15" /></td>
											<td class="alignC">
											<select id="IsLiveWith" name="IsLiveWith" class="w60" title="동거유무">
												<option value="">:유/무:</option>
												<option value="1">유</option>
												<option value="0">무</option>
											</select>
											</td>
										</tr>
									</tbody>
								</table>
							</div>
							<div class="fl_r">
								<span class="btns btn_b"><a href="javascript:addDetailInfo('Fam')">추가</a></span><span class="btns btn_b"><a href="javascript:delDetailInfo('Fam',0)">삭제</a></span>
							</div>
						</div>
					</div> -->
					<div class="section">
						<h2>자기소개서</h2>
						<div class="txtArea">
							<textarea id="Intro" name="Intro" title="자기소개서" rows="8" cols="40" readonly><%=Intro%></textarea>
						</div>
					</div>
					<div class="section">
						<h2>관리자 메모</h2>
						<div class="memoArea">
							<span class="tit_memo"><%=Session("AName")%></span>
							<input type="text" id="Memo" name="Memo" title="메모" value="" class="w560" maxlength="60" />
							<span class="btns btn_b"><a href="javascript:addDetailInfo('Memo')">등록하기</a></span>
						</div>
						<div id="divMemoInfo" class="tblArea">
							<table>
								<colgroup>
									<col class="w80" />
									<col />
									<col class="w150" />
									<col class="w60" />
								</colgroup>
								<thead>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr>
										<th>작성자</th>
										<th>내용</th>
										<th>작성일시</th>
										<th>삭제</th>
									</tr>
								</thead>
								<tbody>
									<% 	
									If IsArray(arrMemo) Then 	
										' UsrMSeq(0), AdmName(1), Memo(2), RegDate(3)
										For i = 0 To UBound(arrMemo, 2) 
									%>
									<tr>
										<td><%=arrMemo(1,i)%></td>
										<td class="alignL"><%=arrMemo(2,i)%></td>
										<td class="alignL"><%=arrMemo(3,i)%></td>
										<td><a href="javascript:delDetailInfo('Memo',<%=arrMemo(0,i)%>)"><img src="/admin/images/btn_del.jpg" alt="삭제" /></a></td>
									</tr>
									<% 
										Next 
									Else
									%>
									<tr>
										<td colspan="4">등록된 메모가 없습니다.</td>
									</tr>
									<%
									End If
									%>
								</tbody>
							</table>
						</div>
					</div>
					<div class="section status_a">
						<h2>지원현황</h2>
						<div class="tblArea">
							<table class="tblForm">
								<colgroup>
									<col class="w80" />
									<col class="w135" />
									<col class="w80" />
									<col />
									<col class="w80" />
									<col class="w45" />
									<col class="w80" />
									<col class="w120" />
								</colgroup>

								<tbody>
									<% 	
									If IsArray(arrApply) Then 	
										' PrjName(0), Title(1), AppBDate(2), AppEDate(3), RegDate(4), AdmName(5), CdVal(6), RecrPart(7), RecrPart2(8), CancelDate(9), RecrSeq(10), CntQues(11)
										For i = 0 To UBound(arrApply, 2) 
										
											strAppCStyle = "" : strAppCMsg = ""

											If Not IsNull(arrApply(9,i)) Then
												strAppCStyle = " class=""cancel"""
												strAppCMsg = " / 지원취소"
											End If
									%>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr<%=strAppCStyle%>>
										<th>프로젝트명</th>
										<td><%=arrApply(0,i)%></td>
										<th>공고제목</th>
										<td colspan="3"><%=arrApply(1,i)%></td>
										<th>담당자</th>
										<td><%=arrApply(5,i)%></td>
									</tr>
									<tr<%=strAppCStyle%>>
										<th>게재기간</th>
										<td><%=FN_SetDateTimeFormat(arrApply(2,i),"YYYY.MM.DD")%> ~ <%=FN_SetDateTimeFormat(arrApply(3,i),"YYYY.MM.DD")%></td>
										<th>상태값</th>
										<td><%=arrApply(6,i)%><%=strAppCMsg%></td>
										<!-- <th>사전질문</th>
										<td>
											<% If arrApply(11,i) > 0 And IsNull(arrApply(9,i)) Then %>
											<span class="btns btn_b" style="margin:auto;"><a href="javascript:goUserQuestion('<%=arrApply(10,i)%>');"  style="padding:0 2px 0 1px; color:#ffffff;">수정</a></span>
											<% End If %>
										</td> -->
										<th>의뢰회차</th>
										<td>
											<%=arrApply(12,i)%>
										</td>
										<th>지원일</th>
										<td><%=FN_SetDateTimeFormat(arrApply(4,i),"YYYY.MM.DD")%></td>
									</tr>
									<% 
										Next 
									Else
									%>
									<tr>
										<td colspan="8">지원 내역이 없습니다.</td>
									</tr>
									<%
									End If
									%>
								</tbody>
							</table>
						</div>
						<!-- div class="regArea">
							<%
							'PartnerSql = "Execute uspPartner_Info @Flag = 'CDD1'" 
							'response.write " ": Call SB_MakeHTMLSelect(objDBCon, PartnerSql, "selPartnerD1_2", "ROOT", "", "", "", " class=""w100""", False)
							%>
							<span id="divPartner_2"></span>
							<%
							'RecruitSql = "Execute uspRecruit_Info @Flag = 'ASLST'" 
							'response.write " ": Call SB_MakeHTMLSelect(objDBCon, RecruitSql, "Recruit_2", "프로젝트 (채용공고) 선택", "", "", "", " class=""w340""", False)
							%>
							<span class="btns btn_b"><a href="javascript:;" id="btnRecr_2">채용공고 등록</a></span>
						</div //-->
						<div class="btnArea">
							<div class="fl_r">
									<span class="btns btn_g"><a href="mem_list.asp?<%=Replace(Filter,"[^]","&")%>">목록보기</a>
									</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a>
									</span><span class="btns btn_b"><a href="javascript:;" id="btnSubmit">등록하기</a></span>
							</div>
						</div>
					</div>
					
					<% End If %>
					</form>
					<form id="frmDown" name="frmDown">
						<input type="hidden" id="dnSeq" name="dnSeq">
						<input type="hidden" id="dnFld" name="dnFld">
						<input type="hidden" id="dnFName" name="dnFName">
						<input type="hidden" id="dnFlag" name="dnFlag">
					</form>

				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>

</body>
</html>