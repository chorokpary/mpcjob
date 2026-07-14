<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : job_form.asp / 채용공고 등록, 수정, 재등록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim PrjSeq, Title, IsMain, IsTop, ChargeTask, TaskType
	Dim Part, RecrPart, RecrPart2, Career, Pay, HireType, LocalSido, LocalGugun, Times
	Dim AppBDate, AppEDate, MapPath, LimitEdu, LimitCareer, LimitBAge, LimitEAge, LimitGender
	Dim ChargeCode, ChargePhone, ChargeEmail, Contents, IsView
	Dim CntMain, CntTop
	
	Dim sViewFlag, sBDate, sEDate, sWord, ListSize
	
	Dim kFields, arrLen
	Dim arrQuestion
	Dim Token : Token = FN_DatetimeToStamp("") & "_" & Session.SessionID & "_" & Session("ASeq")

	Dim i, yy, mm

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)

	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		Flag = CONF_FLAG

		sViewFlag = FN_Req("sViewFlag","")
		sBDate = FN_Req("sBDate","")
		sEDate = FN_Req("sEDate","")
		sWord = FN_Req("sWord","")
		ListSize = FN_Req("ListSize","")
		
		Call getDispCnt
		If Seq <> "0" Then
			Call getData
		End If
		
		If Flag = "ADD" Then
			ChargePhone = "02-3432-6100"
			ChargeEmail = "mpc@mpcjob.co.kr"
		ElseIf Flag = "REF" Then
			AppBDate = Date()
			AppEDate = ""
		End If
		
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			.Parameters("@Token").Value = Token
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PrjSeq 		= Rs("PrjSeq")
				Title 		= FN_InputVal(Rs("Title"))
				IsMain 		= Rs("IsMain")
				IsTop 		= Rs("IsTop")
				ChargeTask 	= FN_InputVal(Rs("ChargeTask"))
				TaskType 	= FN_InputVal(Rs("TaskType"))
				Part 		= Rs("Part")
				RecrPart 	= Rs("RecrPart")
				RecrPart2 	= Rs("RecrPart2")
				Career 		= Rs("Career")
				Pay 		= FN_InputVal(Rs("Pay"))
				HireType 	= Rs("HireType")
				LocalSido 	= Rs("LocalSido")
				LocalGugun 	= Rs("LocalGugun")
				AppBDate 	= FN_SetDateTimeFormat(Rs("AppBDate"),"YYYY-MM-DD") 
				AppEDate 	= FN_SetDateTimeFormat(Rs("AppEDate"),"YYYY-MM-DD") 
				MapPath 	= Rs("MapPath")
				LimitEdu 	= Rs("LimitEdu")
				LimitCareer = Rs("LimitCareer")
				LimitBAge 	= Rs("LimitBAge")
				LimitEAge 	= Rs("LimitEAge")
				LimitGender = Rs("LimitGender")
				ChargeCode 	= Rs("ChargeCode")
				ChargePhone = FN_InputVal(Rs("ChargePhone"))
				ChargeEmail = FN_InputVal(Rs("ChargeEmail"))
				Contents 	= Rs("Contents")
				IsView		= Rs("IsView")
				Times		= ""
			End If

			Set Rs = Rs.NextRecordSet()
				
			'arrDataNum 	= UBound(arrData, 2)
			If Not Rs.Eof Then
				' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4)
				kFields 	= Array("QSeq","QCode","Sort","Question","IsDisp")
				arrQuestion 	= Rs.GetRows(,,kFields)
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub

	Sub getDispCnt()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "CNT_DISP"
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				CntMain 	= Rs("CntMain")
				CntTop 		= Rs("CntTop")
			Else
				CntMain = 0
				CntTop = 0
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
		If Not IsNumeric(CntMain) Then CntMain = 0
		If Not IsNumeric(CntTop) Then CntTop = 0
		
	End Sub
	
%>
	<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
	<script type="text/javascript" src="/common/Dev/SmartEditor/js/HuskyEZCreator.js" charset="utf-8"></script>
	<script type="text/javascript">

		var oEditors = [];
		var RecrPart2 = "<%=RecrPart2%>";

		function getLocalGugun(){
			var url = "/common/dev/ajax/make_select_gugun.asp";
			var param = "SIDO=" + $(this).val() + "&TargetID=LocalGugun&Css=w135";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divGugun").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function getRecrPart2(){
			var url = "/common/dev/ajax/make_select_cd_2depth.asp";
			var param = "Flag=RecrPart&Parent=" + $("#RecrPart").val() + "&TargetID=RecrPart2&Def=" + RecrPart2 + "&Css=w135";
			RecrPart2 = "";
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(content){ $("#divRecrPart2").html(content) }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function chkForm(){
		
			$("#frmInfo").find("#Flag").val("<%=Flag%>");

			if (!$("#Title").val()){
				alert("채용 제목을 입력해주세요.");
				$("#Title").focus();
				return;
			} 

			if (!$("#AppBDate").val()){
				alert("전형일정을 입력해주세요.");
				$("#AppBDate").focus();
				return;
		    } else if(!IsDate($("#AppBDate").val())){
				alert("전형일정은 [년도(4자리)-월(2자리)-일(2자리)] 형식으로 입력해주세요.");
		      	$("#AppBDate").focus();
				return;
			} 

			if (!$("#AppEDate").val()){
				alert("전형일정을 입력해주세요.");
				$("#AppEDate").focus();
				return;
			} else if(!IsDate($("#AppEDate").val())){
				alert("전형일정은 [년도(4자리)-월(2자리)-일(2자리)] 형식으로 입력해주세요.");
		      	$("#AppEDate").focus();
				return;
			}

			if ($("#AppBDate").val() && $("#AppEDate").val()){
				if ($("#AppBDate").val() > $("#AppEDate").val()) {
					alert("전형일정은 시작일이 마감일보다 늦은 날짜일 수 없습니다.");
					$("#AppBDate").focus();
					return;
				}
			} 
			
			<%' If Not IsMain Then %>
			if ($("#IsMain").attr("checked") && <%=CntMain%> >= 24){
				alert("이미 24개의 이달의 추천공고가 등록되어 있습니다.");
				return;
			}
			<%' End If %>

			<%' If Not IsTop Then %>
			if ($("#IsTop").attr("checked") && <%=CntTop%> >= 4){
				alert("이미 4개의 스페셜 HOT이 등록되어 있습니다.");
				return;
			}
			<%' End If %>
			
			if ($("#LimitBAge").val() != "0" && $("#LimitEAge").val() != "0"){
				if ($("#LimitBAge").val() > $("#LimitEAge").val()) {
					alert("연령제한은 시작년도가 마감년도보다 클 수 없습니다.");
					$("#LimitBAge").focus();
					return;
				}
			}
			
			oEditors.getById["Contents"].exec("UPDATE_CONTENTS_FIELD", []);
			
			$("#frmInfo").attr({action:"https://www.mpcjob.co.kr/admin/recruit/job_proc.asp", method:"post", target:"_self", enctype:"multipart/form-data"}).submit();
		}

		function chkIDEnter() {
			if (window.event.keyCode == 13) {
				getSearchAdm()
			}
		}
		
		function getSelectAdm(Flag){
			var newSeq = "";
			
			if (Flag == "ADD"){
				if (cntChecked("#divSearchAdm")==0){
					alert("추가할 관리자를 선택해 주세요.");
					return;
				}else{
					newSeq = getCheckedVal("SearchAdmSeq");
				}
			} else if (Flag =="DEL"){
				if (cntChecked("#divSelectAdm")==0){
					alert("삭제할 관리자를 선택해 주세요.");
					return;
				}else{
					newSeq = getCheckedVal("SelAdmSeq");
				}
			}
			
			var url = "job_adm_ajax.asp";
			var param = "Flag=" + Flag + "&RecrSeq=<%=Seq%>&AdmSeq=" + $("#AdmSeq").val() + "&NewSeq=" + newSeq;

			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(contents){  
					$("#divSelectAdm").html(contents);  
					$("#AdmSeq").val(getCheckboxAllVal("SelAdmSeq"));
					if ($("#SearchKey").val()){
						getSearchAdm();
					}
				}
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function getSearchAdm(){
			if (!$("#SearchKey").val()) {
				alert("검색어를 입력해 주세요.");
				$("#SearchKey").focus();
				return;
			}
		
			var url = "job_adm_ajax.asp";
			var param = "Flag=SEARCH&AdmSeq=" + $("#AdmSeq").val() + "&SearchKey=" + $("#SearchKey").val();
			
			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(contents){  $("#divSearchAdm").html(contents); }
				, error: function(xhr, status, error) {alert("[Rec] " + error);}
			});
		}

		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}

		function addDetailInfo(Flag, DtSize){
			var target;
			if (Flag=="Que") {
				target = (DtSize)?"#QueSort_"+DtSize:"#QueSort";
				if (!$(target).val()){
					alert("노출순서를 입력해주세요.");
					$(target).focus();
					return;
				} 
				if (!IsNumerics($(target).val())){
					alert("노출순서는 숫자만 입력할 수 있습니다.");
					$(target).focus();
					return;
				} 
				target = (DtSize)?"#Question_"+DtSize:"#Question";
				if (!$(target).val()){
					alert("질문을 입력해주세요.");
					$(target).focus();
					return;
				}
			}

			$("#frmInfo").find("#Flag").val(Flag.toUpperCase());
			
			if (DtSize==0)
				$("#frmInfo").find("#FlagSub").val("ADD");
			else
				$("#frmInfo").find("#FlagSub").val("EDIT");

			var url = "https://www.mpcjob.co.kr/admin/recruit/job_proc.asp";
			var param = $("#frmInfo").serialize();
			param += "&IsAjax=1&DtSize="+DtSize ;
			
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
				, error: function(xhr, status, error) {alert("[process] " + error);}
			});
		}

		function getDetailInfo(Flag){
			var url = "job_question_ajax.asp";
			var param = "Flag=" + Flag.toUpperCase() + "&Seq=<%=Seq%>&Token=<%=Token%>" ;

			$.ajax({
				type: "POST"
				, url: url
				, data: param
				, success: function(contents){
					// $("#div" + Flag + "Info").html(contents); 
					$("#div" + Flag + "Info .tblForm tr.list").remove();
					$("#div" + Flag + "Info .tblForm tr.title").after(contents);
					$("#div" + Flag + "Info .tblForm tr.last input:not([type=radio])").val("");
				}, error: function(xhr, status, error) {alert("[road] " + error);}
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
						DelSeq = getCheckedVal(Flag + "Seq");
					}else{
						return;
					}
				}
			}

			var url = "https://www.mpcjob.co.kr/admin/recruit/job_proc.asp";
			var param = "IsAjax=1&Flag=" + Flag.toUpperCase() + "&FlagSub=DEL&Seq=<%=Seq%>&QueSeq=" + DelSeq+"&Token=<%=Token%>";
			
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
				
		function doPreview(){
			$("#PrjName").val($("#PrjSeq").find("option:selected").text());
			$("#ChargeName").val($("#ChargeCode").find("option:selected").text());

			if ($("#DelMapPath").val()){
				$("#TmpImg").val("");
			}

			if ($("#MapPath").val().length > 0) {
				$("#TmpImg").val("");
			}
			
			oEditors.getById["Contents"].exec("UPDATE_CONTENTS_FIELD", []);
			
			$("#frmInfo").attr({action:"/recruit/recruit_view.asp", method:"post", target:"popPreview", enctype:"application/x-www-form-urlencoded"}).submit();
		}
		
		function goList(){
			$("#frmInfo").attr({action:"job_skin_<%=Fn_SetDefault(IsView,True,"list","elist")%>.asp", method:"post", enctype:"application/x-www-form-urlencoded"}).submit();
		}

		function initPage(){
			getSelectAdm("INIT");
			getRecrPart2();
		}
		
		function doReset(){
			$("#frmInfo")[0].reset();
			
			var content = $("#Contents_org").val();
			oEditors.getById["Contents"].exec("SET_IR", [content]);

			initPage();
		}
		
		
		$(function() {
		    $("#LocalSido").change(getLocalGugun);
		    $("#RecrPart").change(getRecrPart2);
		    
			$("#btnSearchAdmin").click(getSearchAdm);
			$("#btnAdmAdd").click(function(){ getSelectAdm('ADD') });
			$("#btnAdmDel").click(function(){ getSelectAdm('DEL') });
			
			$("#btnList").click(goList);
			$("#btnPreview").click(doPreview);
			$("#btnCancel").click(doReset);
			$("#btnSubmit").click(chkForm);

			$("#SearchKey").keydown(chkIDEnter);

			$("input.num").keypress(function(e){ return OnlyNumber(e,false); });
			$("input.num").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
			
			$("#AppBDate").datepicker({
				date: $("#AppBDate").val()
				, current: $("#AppBDate").val()
			});
	
			$("#AppEDate").datepicker({
				date: $("#AppEDate").val()
				, current: $("#AppEDate").val()
			});

			// ### WYSIWYG
			nhn.husky.EZCreator.createInIFrame({
				oAppRef: oEditors,
				elPlaceHolder: "Contents",
				sSkinURI: "/common/Dev/SmartEditor/SmartEditor2Skin.html",	
				fCreator: "createSEditor2"
			});
		})
		
		$(window).load(initPage);
	</script>
	
					<h2>모집요강 <% If Flag = "MOD" Then %><a href="/recruit/recruit_view.asp?Seq=<%=Seq%>" target="_blank">(http://www.mpcjob.co.kr/recruit/recruit_view.asp?Seq=<%=Seq%>)</a><% End If %></h2>

					<form id="frmInfo" name="frmInfo" action="https://www.mpcjob.co.kr/admin/recruit/job_proc.asp" method="post" >
					<input type="hidden" id="Token" name="Token" value="<%=Token%>">
					<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
					<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
					<input type="hidden" id="PageFlag" name="PageFlag" value="ADM">
					<input type="hidden" id="AdmSeq" name="AdmSeq">
					<input type="hidden" id="FlagSub" name="FlagSub">

					<!-- Page //-->
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="sViewFlag" name="sViewFlag" value="<%=sViewFlag%>">
					<input type="hidden" id="sBDate" name="sBDate" value="<%=sBDate%>">
					<input type="hidden" id="sEDate" name="sEDate" value="<%=sEDate%>">
					<input type="hidden" id="sWord" name="sWord" value="<%=sWord%>">
					<input type="hidden" id="ListSize" name="ListSize" value="<%=ListSize%>">
					<input type="hidden" id="IsView" name="IsView" value="<%=IsView%>">
					

					<!-- 미리보기용 //-->
					<input type="hidden" id="ChargeName" name="ChargeName">
					<input type="hidden" id="TmpImg" name="TmpImg" value="<%=MapPath%>">
					<input type="hidden" name="PrjName" id="PrjName">

					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="PrjSeq">프로젝트명</label></th>
									<td>
										<%
										Dim ProjectSql : ProjectSql = "Execute uspProject_Info @Flag = 'ASLST'" 
										Response.Write " ": Call SB_MakeHTMLSelect(objDBCon, ProjectSql, "PrjSeq", "", "", PrjSeq, "", " class=""w275""", False)
										%>
									</td>
									<th><label for="Part">분야</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "Part", "Part", "", "", Part, "등록된 분야가 없습니다.", " class=""w275"" ", False)%>
									</td>
								</tr>
								<tr>
									<th><label for="AppBDate">전형일정</label></th>
									<td colspan="3">
										<input type="text" id="AppBDate" name="AppBDate" value="<%=AppBDate%>" title="전형시작일" style="margin-right:5px;" maxlength="10" /> ~
										<input type="text" id="AppEDate" name="AppEDate" value="<%=AppEDate%>" title="전형마감일" style="margin-right:5px;" maxlength="10"  />
									</td>
								</tr>
								<tr>
									<th><label for="IsTop">공고구분</label></th>
									<td>
										<input type="checkbox" id="IsTop" name="IsTop" title="스페셜 HOT" value="1"<%=Fn_SetDefault(IsTop,True," checked","")%>> <label for="IsTop">스페셜 HOT</label>
										<input type="checkbox" id="IsMain" name="IsMain" title="이달의 추천공고" value="1"<%=Fn_SetDefault(IsMain,True," checked","")%>> <label for="IsMain">이달의 추천공고</label>
									</td>
									<th><label for="TaskType">업무형태</label></th>
									<td>
										<input type="text" id="TaskType" name="TaskType" value="<%=TaskType%>" title="업무형태" class="w270" maxlength="50" />
									</td>
								</tr>
								<tr>
									<th><label for="Title">채용제목</label></th>
									<td colspan="3">
										<input type="text" id="Title" name="Title" value="<%=Title%>" title="채용제목" class="w660" maxlength="50" /> 
									</td>
								</tr>
								<tr>
									<th><label for="ChargeTask">담당업무</label></th>
									<td colspan="3">
										<input type="text" id="ChargeTask" name="ChargeTask" value="<%=ChargeTask%>" title="담당업무" class="w660" maxlength="15" /> 
									</td>
								</tr>
								<tr>
									<th><label for="RecrPart">채용분야</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "RecrPart", "RecrPart", "", "", RecrPart, "등록된 분야가 없습니다.", " class=""w135"" ", False)%>
										<span id="divRecrPart2"></span>
									</td>
									<th><label for="Career">채용종류</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "Career", "Career", "", "", Career, "등록된 종류가 없습니다.", " class=""w275"" ", False)%>
									</td>
								</tr>
								<tr>
									<th><label for="Pay">급여조건</label></th>
									<td>
										<input type="text" id="Pay" name="Pay" value="<%=Pay%>" title="급여조건" class="w270" maxlength="15" /> 
									</td>
									<th><label for="LocalSido">근무지역</label></th>
									<td>
										<%
										Dim SidoSql : SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "LocalSido", "지역", "", LocalSido, "", " class=""w135"" ", False)
										%>
										<span id="divGugun">
										<%
										Dim GugunSql : GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '" & LocalSido & "' GROUP BY GUGUN ORDER BY GUGUN" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, GugunSql, "LocalGugun", "지역중분류", "", LocalGugun, "", " class=""w135"" ", False)
										%>
										</span>
									</td>
								</tr>
								<tr>
									<th><label for="HireType">고용형태</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "HireType", "HireType", "", "", HireType, "등록된 고용형태가 없습니다.", " class=""w275"" ", False)%>
									</td>
									<th><label for="Times">의뢰회차</label></th>
									<td>
										<input type="text" id="Times" name="tiems" value="<%=Times%>" title="의뢰회자" class="w270" maxlength="15" /> 
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>자격요건</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="LimitEdu">학력</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "LimitEdu", "LimitEdu", "", "", LimitEdu, "등록된 학력제한이 없습니다.", " class=""w275"" ", False)%>
									</td>
									<th><label for="LimitBAge">연령</label></th>
									<td>
										<select id="LimitBAge" name="LimitBAge" class="w65">
											<option value="0">--</option>
											<% For yy = Year(Date()) TO GLOBAL_SEL_YEAR - 20 Step -1  %>
											<option value="<%=yy%>"<%=Fn_SetDefault(yy,LimitBAge," selected","")%>><%=yy%></option>
											<% Next %>
										</select>
										년생 ~
										<select id="LimitEAge" name="LimitEAge" class="w65">
											<option value="0">--</option>
											<% For yy = Year(Date()) TO GLOBAL_SEL_YEAR - 20 Step -1  %>
											<option value="<%=yy%>"<%=Fn_SetDefault(yy,LimitEAge," selected","")%>><%=yy%></option>
											<% Next %>
										</select>
										년생 / <input type="checkbox" name="LimitAgeNone" id="LimitAgeNone" value="1" <%=Fn_SetDefault((LimitBAge = "0" And LimitEAge = "0"),True," checked","")%> /><label for="">무관</label>
									</td>
								</tr>
								<tr>
									<th><label for="LimitCareer">경력</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "LimitCareer", "LimitCareer", "", "", LimitCareer, "등록된 경력제한이 없습니다.", " class=""w275"" ", False)%>
									</td>
									<th><label for="LimitGender">성별</label></th>
									<td>
										<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "LimitGender", "LimitGender", "", "", LimitGender, "등록된 성별제한이 없습니다.", " class=""w275"" ", False)%>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>문의 및 담당자 정보</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w85" />
								<col />
								<col class="w85" />
								<col />
								<col class="w85" />
								<col />
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
									<th><label for="ChargePhone">연락처</label></th>
									<td>
										<input type="text" id="ChargePhone" name="ChargePhone" value="<%If ChargePhone <> "" And (Flag = "MOD" Or Flag = "REF") Then Response.write oSeed.Decrypt(ChargePhone) Else Response.write ChargePhone End If %>" title="담당자 연락처" class="w155" maxlength="15" style="ime-mode:disabled;" /> 
									</td>
									<th><label for="ChargeEmail">이메일</label></th>
									<td>
										<input type="text" id="ChargeEmail" name="ChargeEmail" value="<%=ChargeEmail%>" title="담당자 이메일" class="w155" maxlength="50" style="ime-mode:disabled;" /> 
									</td>
									<th><label for="ChargeCode">담당자</label></th>
									<td>
										<%
										Dim AdmSql : AdmSql = "Execute uspAdm_Info @Flag = 'LV1'" 
										Response.Write " ": Call SB_MakeHTMLSelect(objDBCon, AdmSql, "ChargeCode", "", "", ChargeCode, "", " class=""w165""", False)
										%>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>상세내용 등록</h2>
					<div class="tblArea">
						<textarea name="Contents" id="Contents" class="w660" style="display:none;"><%=Contents%></textarea>
						<textarea name="Contents_org" id="Contents_org" class="w660" style="display:none;"><%=Contents%></textarea>
					</div>
					<h2>오시는길 등록</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w85" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									<th><label for="MapPath">이미지등록</label></th>
									<td>
										<% If Not FN_isBlank(MapPath) Then %>
										<p id="pMapPath" class="filename"><!--img src="<%=MapPath%>" width="116" height="144" //--><%=MID(MapPath, InStrRev(MapPath,"/") + 1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pMapPath','DelMapPath','<%=MapPath%>')">삭제</a></span></p>
										<% End If %>

										<input type="file" id="MapPath" name="MapPath" title="이미지첨부" value="" class="w660" size="95" />
										<input type="hidden" id="DelMapPath" name="DelMapPath" value="">
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					<h2>사전질문</h2>
					<div id="divQuestion" class="tblArea">
						<div id="divQueInfo">
							<%
							If IsArray(arrQuestion) Then
								arrLen = UBound(arrQuestion, 2)
							Else
								arrLen = -1
							End If
							%>
							<table class="tblForm">
								<colgroup>
									<col class="w70" />
									<col />
									<col class="w120" />
									<col class="w65" />
									<col class="w65" />
								</colgroup>
								<tbody>
									<tr class="design">
										<td></td>
										<td></td>
										<td></td>
										<td></td>
										<td></td>
									</tr>
									<tr class="title">
										<th>노출순서<br/>(오름차순)</th>
										<th>질문</th>
										<th>노출여부</th>
										<th colspan="2"></th>
									</tr>
									<% 		
									' ' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4)
									Dim Num
									For i = 0 To arrLen 
										Num = i + 1
									%>
									<tr class="list">
										<td class="alignC"><input type="text" name="QueSort_<%=Num%>" id="QueSort_<%=Num%>" title="노출순서" value="<%=FN_InputVal(arrQuestion(2,i))%>" class="w50 alignC num" maxlength="3" /></td>
										<td class="alignC"><input type="text" name="Question_<%=Num%>" id="Question_<%=Num%>" value="<%=FN_InputVal(arrQuestion(3,i))%>" class="w530" maxlength="200"></td>
										<td class="alignC">
											<input type="radio" name="QueIsDisp_<%=Num%>" id="QueIsDisp_<%=Num%>_Y" value="1" <%=Fn_SetDefault(arrQuestion(4,i), "True", " checked", "")%>><label for="QueIsDisp_<%=Num%>_Y">노출</label>
											<input type="radio" name="QueIsDisp_<%=Num%>" id="QueIsDisp_<%=Num%>_N" value="0" <%=Fn_SetDefault(arrQuestion(4,i), "False", " checked", "")%>> <label for="QueIsDisp_<%=Num%>_N">숨김</label>
										</td>
										<td>
											<span class="btns btn_gr"><a href="javascript:delDetailInfo('Que',<%=arrQuestion(0,i)%>)">삭제</a></span>
											<input type="hidden" id="QueSeq_<%=Num%>" name="QueSeq_<%=Num%>" value="<%=arrQuestion(0,i)%>" />
										</td>
										<% If i = 0 Then %>
										<td rowspan="<%=arrLen+1%>">
											<span class="btns btn_b"><a href="javascript:addDetailInfo('Que',<%=arrLen+1%>)">수정</a></span>
										</td>
										<% End If %>
									</tr>
									<% Next %>
									<tr class="last">
										<td class="alignC"><input type="text" name="QueSort" id="QueSort" title="노출순서" value="" class="w50 alignC num" maxlength="3" /></td>
										<td class="alignC"><input type="text" name="Question" id="Question" class="w530" maxlength="200"></td>
										<td class="alignC">
											<input type="radio" name="QueIsDisp" id="QueIsDisp_Y" value="1" checked><label for="QueIsDisp_Y">노출</label>
											<input type="radio" name="QueIsDisp" id="QueIsDisp_N" value="0"> <label for="QueIsDisp_N">숨김</label>
										</td>
										<td colspan="2">
											<span class="btns btn_b"><a href="javascript:addDetailInfo('Que',0)">추가</a></span>
										</td>
									</tr>
								</tbody>
							</table>
						</div>
						<p class="note_c">* 추가 및 수정 후 등록하기 버튼을 클릭해야 저장됩니다. 새로고침 할 경우 작성한 내용이 사라집니다.</p>
						<!-- div class="fl_r mt5">
							<span class="btns btn_b"><a href="javascript:delDetailInfo('Que',0)">선택삭제</a></span>
						</div //-->
					</div>
					<h2>채용공고 심사권한</h2>
					<div class="tblArea">
						<table class="tblForm">
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
								<tr>
									
									<th>이름검색</th>
									<td>
										<input type="text" id="SearchKey" name="SearchKey" value="" title="이름검색" class="w270" /> 
										<span class="btns btn_in"><a href="#btnSearchAdmin" id="btnSearchAdmin">검색</a></span>
									</td>
								</tr>
								<tr>
									
									<th>선택</th>
									<td>
										<div id="divSearchAdm" class="pick">
											<p>검색된 인원명 : 0명</p>
											<div class="innerBox"></div>
										</div>
										<p class="pick_btn">
											<a href="#btnAdmAdd" id="btnAdmAdd" class="btn_reg_add"><img src="/admin/images/btn_reg_add.png" alt="추가" /></a>
											<a href="#btnAdmDel" id="btnAdmDel" class="btn_reg_del"><img src="/admin/images/btn_reg_del.png" alt="삭제" /></a>
										</p>
										<div id="divSelectAdm" class="pick"><!-- 선택된 인원 //--></div>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					</form>
					<div class="btnArea">
						<div class="fl_r">
							<span class="btns btn_g"><a href="#btnList" id="btnList">목록보기</a>
							</span><span class="btns btn_g"><a href="#btnPreview" id="btnPreview">미리보기</a>
							</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a>
							</span><span class="btns btn_b"><a href="#" id="btnSubmit">등록하기</a></span>
						</div>
					</div>
