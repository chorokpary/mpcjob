<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : resume_modify.asp / 이력서 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 2013-12-xx / 강미현 / 4M / 희망분야 > 채용 분야로 통합
'   #002 : 2014-01-07 / 강미현 / 4M / 희망분야 등록항목 삭제
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Flag, Page
	Dim UsrName, Birthday, Gender, Mobile, Email, HopePart, HopePart2, UsrType, Email_1, Email_2
	Dim Zipcode1, Zipcode2, AddrSido, AddrGugun, Addr, AddrDetail, Jumin1, Jumin2
	Dim ImgPath, Attach1_Path, Attach2_Path, Attach3_Path, Attach1_FName, Attach2_FName, Attach3_FName
	Dim RegDate, ModDate, Intro, EduLast, EduType
	Dim kFields, arrLen, i, yy, mm, dd
	Dim arrEdu, arrCareer, arrFam, arrMemo, arrApply, arrEzCareer
	Dim arrCdEzCr
	Dim strEzCareer
	Dim SidoSql, GugunSql
	
	Call Main()
	
	Sub Main()

		Page = FN_Req("Page","1")
		
		If Session("USeq") <> "0" Then
			Flag = "MOD"
			Call getData
			Call getCdData
		Else
			Call SB_ReturnErr("잘못된 경로로 접근하였습니다.", "BACK")
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
			.Parameters("@Seq").Value 	= Session("USeq")
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				UsrName 	= FN_InputVal(Rs("UsrName"))
				Birthday 	= Rs("Birthday")
				Gender 	= Rs("Gender")
				Mobile 	= FN_InputVal(Rs("Mobile"))
				Email 	= FN_InputVal(Rs("Email"))
				
				If InStr(Email,"@") Then 
					Email_1 = Split(Email,"@")(0)
					Email_2 = Split(Email,"@")(1)
				Else
					Email_1 = ""
					Email_2 = ""
				End If 

				HopePart 	= Rs("HopePart")
				HopePart2 = Rs("HopePart2")
				UsrType 	= Rs("UsrType")
				Zipcode1 = Rs("Zipcode1")
				Zipcode2 = Rs("Zipcode2")
				AddrSido 	= FN_InputVal(Rs("AddrSido"))
				AddrGugun = FN_InputVal(Rs("AddrGugun"))
				Addr 	= FN_InputVal(Rs("Addr"))
				AddrDetail = FN_InputVal(Rs("AddrDetail"))
				Jumin1 	= FN_InputVal(Rs("Jumin1"))
				Jumin2 	= FN_InputVal(Rs("Jumin2"))

				If AddrSido = "" Then AddrSido = "서울"
				If AddrGugun = "" Then AddrGugun = "지역중분류"

				ImgPath 	= Rs("ImgPath")
				Attach1_Path = Rs("Attach1_Path")
				Attach1_FName = Rs("Attach1_FName")
				Attach2_Path 	= Rs("Attach2_Path")
				Attach2_FName = Rs("Attach2_FName")
				Attach3_Path 	= Rs("Attach3_Path")
				Attach3_FName = Rs("Attach3_FName")
				
				RegDate = Rs("RegDate")
				ModDate = Rs("ModDate")
				
				Intro = FN_BrToNL(Rs("Intro"))

				EduLast = Rs("EduLast")
				EduType = Rs("EduType")


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
				kFields 	= Array("UsrSeq","CrBYear","CrBMonth","CrEYear","CrEMonth","CompanyName","Department","Position","Task","Reason", "Sort" ,"CompanyYear")
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
				' UsrMSeq(0), AdmName(1), Memo(2)
				kFields 	= Array("UsrMSeq","AdmName","Memo")
				arrMemo 	= Rs.GetRows(,,kFields)
			End If

			Set Rs = Rs.NextRecordSet()

			If Not Rs.Eof Then
				' PrjName(0), Title(1), AppBDate(2), AppEDate(3), RegDate(4), AdmName(5), CdVal(6)
				kFields 	= Array("PrjName","Title","AppBDate","AppEDate","RegDate","AdmName","CdVal")
				arrApply 	= Rs.GetRows(,,kFields)
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

	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>MY PAGE - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		var HopePart2 = "<%=HopePart2%>";

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
	
			$("#frmInfo").find("#Flag").val("<%=Flag%>");

			$("#Email").val($("#email1").val() + "@" + $("#email2").val());

			if (!$("#UsrName").val()){
				alert("이름을 입력해주세요.");
				$("#UsrName").focus();
				return;
			/*
			} else if (!$("#Birthday_Y").val()){
				alert("생년월일을 입력해주세요.");
				$("#Birthday_Y").focus();
				return;
			} else if (!$("#Birthday_M").val()){
				alert("생년월일을 입력해주세요.");
				$("#Birthday_M").focus();
				return;
			} else if (!$("#Birthday_D").val()){
				alert("생년월일을 입력해주세요.");
				$("#Birthday_D").focus();
				return;
		    } else if(!IsDate($("#Birthday_Y").val() + "-" + $("#Birthday_M").val() + "-" + $("#Birthday_D").val() )){
				alert("생년월일이 날짜 범위를 벗어납니다.");
		      	$("#Birthday_Y").focus();
				return;
			*/
			/*
			} else if (!$("#HopePart").val()){
				alert("희망분야를 선택해주세요.");
				$("#HopePart").focus();
				return;
			
			} else if (!$("#Jumin1").val()){
				alert("주민등록번호를 입력해주세요.");
				$("#Jumin1").focus();
				return;
			} else if (!$("#Jumin2").val()){
				alert("주민등록번호를 입력해주세요.");
				$("#Jumin2").focus();
				return;
			
			
			}  else if (!$("#Addr").val()){
				alert("주소를 입력해주세요.");
				$("#Addr").focus();
				return;
			*/
			}  else if (!$("#frmInfo").find("#AddrSido").val()){
				alert("지역을 선택해주세요.");
				$("#frmJoin").find("#AddrSido").focus();
				return;
			}  else if (!$("#frmInfo").find("#AddrGugun").val()){
				alert("지역중분류를 선택해주세요.");
				$("#frmJoin").find("#AddrGugun").focus();
				return;
			} else if (!$("#Email").val()){
				alert("E-mail을 입력해주세요.");
				$("#Email").focus();
				return;
			}
			
		
			$("#frmInfo").attr({action:"resume_proc.asp", method:"post", enctype:"multipart/form-data"}).submit();
		}
		
	
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}
		
		function addDetailInfo(Flag){
			if (Flag=="Edu") {
				if (!$("#SchoolName").val()){
					//alert("학교명을 입력해주세요.");
					//$("#SchoolName").focus();
					//return;
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
			}

			$("#frmInfo").find("#Flag").val(Flag.toUpperCase());
			$("#frmInfo").find("#FlagSub").val("ADD");

			var url = "resume_proc.asp";
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
			var url = "resume_info_ajax.asp";
			var param = "Flag=" + Flag.toUpperCase() ;

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

			var url = "resume_proc.asp";
			var param = "IsAjax=1&Flag=" + Flag.toUpperCase() + "&FlagSub=DEL&Sort=" + DelSeq;
			
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
		
		function doSearchZipcode(){
			var param = "tgSi=AddrSido&tgGu=AddrGugun&tgZip1=Zipcode1&tgZip2=Zipcode2&tgAddr=Addr&tgAddr2=AddrDetail";
			openLayer("zipcode.asp",param, "w528");
		
		}
		


		function getJoinAddrGugun(){
		var url = "/common/dev/ajax/make_select_gugun.asp";
		var param = "SIDO=" + $(this).val() + "&TargetID=AddrGugun&Css=w115";
		
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divGugun").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}


		$(function() {
			$("#AddrSido").change(getJoinAddrGugun);
			$("#HopePart").change(getHopePart2);
			
			$("#btnFindZip").click(function(){ openWin("../popup/zipcode.asp?","popZipcode",0,0,550,300,0); });
			
			$("#Jumin1").keypress(function(e){ return OnlyNumber(e,false); });
			$("#Jumin1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });

			$("#Jumin2").keypress(function(e){ return OnlyNumber(e,false); });
			$("#Jumin2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
			
			$(document).ready(getHopePart2);
			
		})
	</script>
</head>
<body>
	<div id="wrap" class="resumeMod resumeView">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_resumeMng.jpg" alt="이력서관리" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				
				<form id="frmInfo" name="frmInfo" action="mem_proc.asp" method="post" enctype!="multipart/form-data" >
				<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
				<input type="hidden" id="FlagSub" name="FlagSub">

				<div class="section">
					<table class="tblForm">
						<colgroup>
							<col class="w100" />
							<col class="w170" />
							<col class="w90" />
							<col />
						</colgroup>
						<tbody>
							<tr>
								<th><label for="UsrName"><span class="ess">(*)</span> 이름</label></th>
								<td>
									<input type="text" id="UsrName" name="UsrName" title="이름" value="<%=UsrName%>" class="w145" maxlength="25" /> 
								</td>
								<th><label for="Birthday_Y">생년월일</label></th>
								<td>
									<select id="Birthday_Y" name="Birthday_Y" title="생년" class="w50">
										<% For yy = Year(Date()) To GLOBAL_SEL_YEAR Step -1  %>
										<option value="<%=yy%>" <%=Fn_SetDefault(yy,FN_SetDateTimeFormat(BirthDay,"YYYY")," selected","")%>><%=yy%></option>
										<% Next %>
									</select> 년
									<select id="Birthday_M" name="Birthday_M" title="생월" class="w45">
										<% For mm = 1 TO 12 %>
										<option value="<%=Right("0" & mm,2)%>" <%=Fn_SetDefault(Right("0" & mm,2),FN_SetDateTimeFormat(BirthDay,"MM")," selected","")%>><%=Right("0" & mm,2)%></option>
										<% Next %>
									</select> 월
									<select id="Birthday_D" name="Birthday_D" title="생일" class="w45">
										<% For dd = 1 TO 31 %>
										<option value="<%=Right("0" & dd,2)%>" <%=Fn_SetDefault(Right("0" & dd,2),FN_SetDateTimeFormat(BirthDay,"DD")," selected","")%>><%=Right("0" & dd,2)%></option>
										<% Next %>
									</select> 일
								</td>

							</tr>
							<tr>
								<th><label for=""><span class="ess">(*)</span>핸드폰번호</label></th>
								<td><%=Mobile%></td>
								<th><label for="Gender">성별</label></th>
								<td>
									<input type="radio" name="Gender" id="Gender1" value="1" <%=Fn_SetDefault(Gender,"1"," checked","")%> >남
									<input type="radio" name="Gender" id="Gender2" value="2" <%=Fn_SetDefault(Gender,"2"," checked","")%>>여
								</td>
							</tr>
							
							
							<!-- <tr>
								<th><label for="Zipcode1"><span class="ess">(*)</span>우편번호</label></th>
								<td colspan="3">
									<input type="number" id="Zipcode1" name="Zipcode1" title="우편번호" value="<%=Zipcode1%>" maxlength="3" class="w40" readonly />
									-
									<input type="number" id="Zipcode2" name="Zipcode2" title="우편번호" value="<%=Zipcode2%>" maxlength="3" class="w40" readonly />
									<a href="javascript:doSearchZipcode();"><img src="/images/pages/btn_zipcode.jpg" alt="우편번호검색" /></a>
									<input type="hidden" id="AddrSido" name="AddrSido" value="<%=AddrSido%>">
									<input type="hidden" id="AddrGugun" name="AddrGugun" value="<%=AddrGugun%>">
								</td>
							</tr> -->

							<tr>
								<th><label for="Zipcode1"><span class="ess">(*)</span>주소</label></th>
								<td colspan="3">
									
									<%
										SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY CASE SIDO WHEN '"&AddrSido&"' THEN 0 ELSE 1 END ASC, SIDO ASC" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "AddrSido", "", AddrSido, "", "", " class=""w160"" ", False)
									%>

									<span id="divGugun"> 
									<%
									GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '"&AddrSido&"' GROUP BY GUGUN ORDER BY GUGUN" 
									response.write " ": Call SB_MakeHTMLSelect(objDBCon, GugunSql, "AddrGugun", AddrGugun, AddrGugun, "", "", " class=""w220"" ", False)
									%>
									</span>
								</td>
							</tr>

							
							<!-- <tr>
								<th><label for="Addr"><span class="ess">(*)</span>주소</label></th>
								<td colspan="3">
									<input type="text" id="Addr" name="Addr" title="주소" value="<%=Addr%>" class="w640" readonly maxlength="75" />
								</td>
							</tr> -->
							<!-- <tr>
								<th><label for="AddrDetail">상세주소</label></th>
								<td colspan="3">
									<input type="text" id="AddrDetail" name="AddrDetail" title="상세주소" value="<%=AddrDetail%>" class="w640" maxlength="75" />
								</td>
							</tr> -->

							<tr>
								<th><label for="Email"><span class="ess">(*)</span>E-mail</label></th>
								<td colspan="3">
									
									<input type="hidden" id="Email" name="Email" title="E-mail 주소" value="<%=Email%>" /> 
									<input type="text" id="email1" name="email1" class="w114" title="이메일 주소의 앳(@) 이전 정보 입력" maxlength="18" value="<%=Email_1%>"  />
									@
									<input type="text" id="email2" name="email2" class="w144" title="이메일 주소의 앳(@) 이후 정보 직접입력" maxlength="30"  value="<%=Email_2%>" />
									<select id='emailDomains' name='emailDomains' title="이메일 주소의 앳(@) 이후 정보 선택" style="width:160px;"  onchange="document.frmInfo.email2.value=this.value;document.frmInfo.Email.value=document.frmInfo.email1.value+'@'+this.value;">
										<option>직접입력</option>
										<option value=chol.com>chol.com</option>
										<option value=dreamwiz.com>dreamwiz.com</option>
										<option value=empal.com>empal.com</option>
										<option value=freechal.com>freechal.com</option>
										<option value=gmail.com>gmail.com</option>
										<option value=hanafos.com>hanafos.com</option>
										<option value=hanmail.net>hanmail.net</option>
										<option value=hanmir.com>hanmir.com</option>
										<option value=hitel.net>hitel.net</option>
										<option value=hotmail.com>hotmail.com</option>
										<option value=korea.com>korea.com</option>
										<option value=lycos.co.kr>lycos.co.kr</option>
										<option value=nate.com>nate.com</option>
										<option value=natian.com>natian.com</option>
										<option value=naver.com>naver.com</option>
										<option value=netian.com>netian.com</option>
										<option value=paran.com>paran.com</option>
										<option value=sayclub.com>sayclub.com</option>
										<option value=yahoo.co.kr>yahoo.co.kr</option>
										<option value=yahoo.com>yahoo.com</option>
									</select>

									<!-- <input type="email" id="Email" name="Email" title="E-mail 주소" value="<%=Email%>" class="w285" maxlength="100" style="ime-mode:disabled;" />  -->
								</td>
							</tr>
						</tbody>
					</table>
					<p class="note"><span class="ess">(*)</span>는 필수 입력항목입니다.</p>
				</div>
				
				<h3><img src="/images/pages/stit_resumeMod_02.jpg" alt="학력사항" /></h3>
				<div class="section">
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
	
						<% If arrLen > -1 Then %>
						<!-- <table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr>
									<th>최종학력</th>
									<td class="alignL"> <strong><% If arrLen > -1 Then %> <%=strFinEdu%> <% End If %></strong></th>
								</tr>
							</tbody>
						</table> -->
						<% End If %>
						
						<table class="tblForm alignC">
							<colgroup>
								<col class="w150" />
								<col />
							</colgroup>
							<tbody>
								
								<!-- <tr>
									<th></th>
									<th>입학년/월</th>
									<th>졸업년/월</th>
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
									<td><input type="checkbox" id="EduSort_<%=arrEdu(9,i)%>" name="EduSort" value="<%=arrEdu(9,i)%>" title="선택" /></td>
									<td>
										<select id="EduLast_<%=i%>" name="EduLast" title="최종학력" class="w120">
											<option value="0">중학교</option>
											<option value="1">고등학교</option>
											<option value="2">대학(2,3년)</option>
											<option value="3">대학교(4년)</option>
											<option value="4">대학원</option>
										</select>
									</td>
									<td class="alignL">
										&nbsp;<input type="radio" name="EduType" id="EduType1" value="1">졸업
										<input type="radio" name="EduType" id="EduType2" value="2">재학중
										<input type="radio" name="EduType" id="EduType3" value="3">휴학
										<input type="radio" name="EduType" id="EduType4" value="4">중퇴
									</td>
								</tr> -->
								<% Next %>
								<tr>
									<td>
										<select id="EduLast" name="EduLast" title="최종학력" class="w120">
											<option value="0" <%=Fn_SetDefault(EduLast,"0"," selected","")%>>중학교</option>
											<option value="1" <%=Fn_SetDefault(EduLast,"1"," selected","")%>>고등학교</option>
											<option value="2" <%=Fn_SetDefault(EduLast,"2"," selected","")%>>대학(2,3년)</option>
											<option value="3" <%=Fn_SetDefault(EduLast,"3"," selected","")%>>대학교(4년)</option>
											<option value="4" <%=Fn_SetDefault(EduLast,"4"," selected","")%>>대학원</option>
										</select>
									</td>
									<td class="alignL">
										&nbsp;<input type="radio" name="EduType" id="EduType1" value="1" <%=Fn_SetDefault(EduType,"1"," checked","")%> >졸업
										<input type="radio" name="EduType" id="EduType2" value="2" <%=Fn_SetDefault(EduType,"2"," checked","")%>>재학중
										<input type="radio" name="EduType" id="EduType3" value="3" <%=Fn_SetDefault(EduType,"3"," checked","")%>>휴학
										<input type="radio" name="EduType" id="EduType4" value="4" <%=Fn_SetDefault(EduType,"4"," checked","")%>>중퇴
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					 <!-- <div class="alignR pt10"> <a href="javascript:addDetailInfo('Edu')"><img src="/images/pages/btn_save.jpg" alt="저장" /></a> <a href="javascript:delDetailInfo('Edu',0)"><img src="/images/pages/btn_del.jpg" alt="삭제" /></a></div> -->
				</div>
				
				
				<h3><img src="/images/pages/stit_resumeMod_03.jpg" alt="경력사항" /></h3>
				<div class="section">
					<div id="divCareerInfo">
						<%
						If IsArray(arrCareer) Then
							arrLen = UBound(arrCareer, 2)
						Else
							arrLen = -1
						End If
						%>
						<table class="tblForm alignC">
							<colgroup>
								<col class="w30" />
								<col class="w100" />
								<col class="w80" />
								<col class="w80" />
								<col class="w200" />
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<!-- <tr>
									<th></th>
									<th>근무기간</th>
									<th>직장명</th>
									<th>직위</th>
									<th>부서명</th>
									<th>담당업무</th>
									<th>퇴직사유</th>
								</tr> -->
								<% 		
								'  UsrSeq(0), CrBYear(1), CrBMonth(2), CrEYear(3), CrEMonth(4), CompanyName(5), Department(6), Position(7), Task(8), Reason(9), Sort(10), CompanyYear(11)
								For i = 0 To arrLen 
								%>
								<tr>
									<td class="alignC"><input type="checkbox" id="CareerSort_<%=arrCareer(10,i)%>" name="CareerSort" value="<%=arrCareer(10,i)%>" title="선택" /></td>
									<td class="alignC">근무개월수</td>
									<td>
										<%
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
										%>
									</td>
									<td class="alignC">회사명</td>
									<td class="alignC"><%=arrCareer(5,i)%></td>
									<td>담당업무</td>
									<td><%=arrCareer(8,i)%></td>
								</tr>
								<% Next %>
								<tr>
									<td class="alignC"></td>
									<td class="alignC">근무개월수</td>
									<td>
										<select id="CompanyYear" name="CompanyYear" title="근무개월수" class="w70">
											<option value="0">1년미만</option>
											<option value="1">1년이상</option>
											<option value="2">2년이상</option>
											<option value="3">3년이상</option>
											<option value="4">4년이상</option>
										</select>
									</td>
									<td class="alignC">회사명</td>
									<td class="alignC"><input type="text" id="CompanyName" name="CompanyName" title="직장명" value="" class="w125" maxlength="20" /></td>
									<td>담당업무</td>
									<td><input type="text" id="Task" name="Task" name="Task" title="업무" value="" class="w125" maxlength="15" /></td>
								</tr>
							</tbody>
						</table>
					</div>
					<div class="alignR pt10"> <a href="javascript:addDetailInfo('Career');"><img src="/images/pages/btn_add.jpg" alt="추가" /></a> <a href="javascript:delDetailInfo('Career',0);"><img src="/images/pages/btn_del.jpg" alt="삭제" /></a></div>
				</div>

				<h3><img src="/images/pages/stit_resumeMod_06.jpg" alt="첨부파일" /></h3>
				<div class="section">
					<table class="tblForm">
						<colgroup>
							<col class="w100" />
							<col class="w170" />
							<col class="w90" />
							<col />
							<col class="w80" />
							<col class="w90" />
						</colgroup>
						<tbody>
							<tr>
								<th><label for="ImgPath">이미지</label></th>
								<td colspan="5">
									<% If Not FN_isBlank(ImgPath) Then %>
									<div id="pImgPath" class="imgs">
										<div class="inner">
											<img src="<%=ImgPath%>" width="114" height="144"> 
										</div>
										<a href="javascript:doDelFile('pImgPath','DelImgPath','<%=ImgPath%>');" class="btn_attach_del"><img src="/images/common/btn_attach_del.jpg" alt="삭제" /></a>
									</div>
									<% Else %>
									<div class="imgs">
										<div class="inner">
											<img src="/images/pages/img_nopics.jpg" alt="" />
										</div>
									</div>
									<% End If %> 
									<input type="file" id="ImgPath" name="ImgPath" title="이미지첨부" value="" class="w650" size="92" /> 
									<input type="hidden" id="DelImgPath" name="DelImgPath" value="">
								</td>
							</tr>
							<tr>
								<th><label for="Attach_1">이력서 1</label></th>
								<td colspan="5">
									<% If Not FN_isBlank(Attach1_Path) Then %>
									<p id="pAttach_1"><%=Attach1_FName%> <a href="javascript:doDelFile('pAttach_1','DelAttach_1','<%=Attach1_Path%>')" class="btn_attach_del"><img src="/images/common/btn_attach_del.jpg" alt="삭제" /></a></p>
									<% End If %>
									<input type="file" id="Attach_1" name="Attach_1" title="첨부파일 1" value="" class="w650" size="92" /> 
									<input type="hidden" id="DelAttach_1" name="DelAttach_1" value="">
								</td>
							</tr>
							<tr>
								<th><label for="Attach_2">이력서 2</label></th>
								<td colspan="5">
									<% If Not FN_isBlank(Attach2_Path) Then %>
									<p id="pAttach_2"><%=Attach2_FName%> <a href="javascript:doDelFile('pAttach_2','DelAttach_2','<%=Attach2_Path%>')" class="btn_attach_del"><img src="/images/common/btn_attach_del.jpg" alt="삭제" /></a></p>
									<% End If %>
									<input type="file" id="Attach_2" name="Attach_2" title="첨부파일 2" value="" class="w650" size="92" /> 
									<input type="hidden" id="DelAttach_2" name="DelAttach_2" value="">
								</td>
							</tr>
							<!-- <tr>
								<th><label for="Attach_3">첨부파일 3</label></th>
								<td colspan="5">
									<% If Not FN_isBlank(Attach3_Path) Then %>
									<p id="pAttach_3"><%=Attach1_FName%> <a href="javascript:doDelFile('pAttach_3','DelAttach_3','<%=Attach3_Path%>')" class="btn_attach_del"><img src="/images/common/btn_attach_del.jpg" alt="삭제" /></a></p>
									<% End If %>
									<input type="file" id="Attach_3" name="Attach_3" title="첨부파일 3" value="" class="w650" size="92" /> 
									<input type="hidden" id="DelAttach_3" name="DelAttach_3" value="">
								</td>
							</tr> -->
						</tbody>
					</table>

				</div>

				<!-- <h3><img src="/images/pages/stit_resumeMod_04.jpg" alt="가족사항" /></h3>
				<div class="section">
					<div id="divFamInfo">
						<%
						If IsArray(arrFam) Then
							arrLen = UBound(arrFam, 2)
						Else
							arrLen = -1
						End If
						%>

						<table class="tblForm alignC">
							<colgroup>
								<col class="w30" />
								<col class="w115" />
								<col class="w155" />
								<col class="w75" />
								<col class="w115" />
								<col />
								<col class="w100" />
							</colgroup>
							<tbody>
								<tr>
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
					
					<div class="alignR pt10"> <a href="javascript:addDetailInfo('Fam');"><img src="/images/pages/btn_add.jpg" alt="추가" /></a> <a href="javascript:delDetailInfo('Fam',0);"><img src="/images/pages/btn_del.jpg" alt="삭제" /></a></div>
				</div> -->
				<h3><img src="/images/pages/stit_resumeMod_05.jpg" alt="자기소개서" /></h3>
				<div class="textAreas">
					<textarea name="Intro" id="Intro" rows="8" cols="40"><%=Intro%></textarea>
				</div>
				<div class="alignR pt10"><a href="javascript:chkForm();"><img src="/images/pages/btn_save.jpg" alt="저장" /></a></div>
				</form>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>