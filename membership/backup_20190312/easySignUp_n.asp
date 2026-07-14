<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
	Dim PageFlag, RecrSeq
	Dim yy, mm
	Dim arrCdEzCr, i
	

	' 로그인시 접근 제한 페이지
	Call SB_NoLoginPage()
	Call Main()
	
	Sub Main()

		PageFlag = FN_Req("PageFlag","JOIN")
		PageFlag = "JOIN"
		RecrSeq = FN_Req("RecrSeq","0")
		
		If PageFlag = "JOIN" Then
			RecrSeq = 0
		End If
		
		Call getCdData
		
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

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>채용정보 - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	
<script type="text/javascript">
	function chkJoinForm(){
		
		$("#frmJoin").find("#Flag").val("ADD");

		$("#Email").val($("#email1").val() + "@" + $("#email2").val());

		if (!$("#frmJoin").find("#UsrName").val()){
			alert("이름을 입력해주세요.");
			$("#frmJoin").find("#UsrName").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_Y").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_Y").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_M").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_M").focus();
			return;
		} else if (!$("#frmJoin").find("#Birthday_D").val()){
			alert("생년월일을 입력해주세요.");
			$("#frmJoin").find("#Birthday_D").focus();
			return;
	    } else if(!IsDate($("#frmJoin").find("#Birthday_Y").val() + "-" + $("#frmJoin").find("#Birthday_M").val() + "-" + $("#frmJoin").find("#Birthday_D").val() )){
			alert("생년월일이 날짜 범위를 벗어납니다.");
	      	$("#frmJoin").find("#Birthday_Y").focus();
			return;			
		} else if (!$("#frmJoin").find("#Mobile_1").val() || $("#frmJoin").find("#Mobile_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_2").val() || $("#frmJoin").find("#Mobile_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_3").val() || $("#frmJoin").find("#Mobile_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_3").focus();
			return;
		} else if (!IsMobile($("#frmJoin").find("#Mobile_1").val() + $("#frmJoin").find("#Mobile_2").val() + $("#frmJoin").find("#Mobile_3").val())){
			alert("핸드폰번호가 바르지 않습니다.");
			$("#frmModify").find("#Mobile_1").focus();
			return;
		
		} else if (!$("#frmJoin").find("#AddrSido").val()){
			alert("지역을 선택해주세요.");
			$("#frmJoin").find("#AddrSido").focus();
			return;
		} else if (!$("#frmJoin").find("#AddrGugun").val()){
			alert("지역중분류를 선택해주세요.");
			$("#frmJoin").find("#AddrGugun").focus();
			return;
		} else if (!$("#frmJoin").find("#Passwd").val()){
			alert("비밀번호를 입력해주세요.");
			$("#frmJoin").find("#Passwd").focus();
			return;
		}else if(!fn_pw_check()){
			return ;
		}else if (!$("#chk_consent1").attr("checked")){
			alert("서비스 약관에 동의하셔야 합니다.");
			return;
		}else if (!$("#chk_consent2").attr("checked")){
			alert("개인정보 수집에 동의하셔야 합니다.");
			return;
		}

			
			
		if($("#frmJoin").find("#ImgPath").val() == "" && $("#frmJoin").find("#Attach_1").val() == "" && $("#frmJoin").find("#Attach_2").val() == "" ){
			$("#PageFlag").val('NOFILE');
		}

		// 핸드폰번호 중복 체크
		doChkMobile();

		//$("#frmJoin").attr({action:"/membership/easySignUp_proc.asp", method:"post", enctype:"multipart/form-data"}).submit();
		
	}
		
	function doChkMobile() {

		var url = "/common/dev/ajax/check_exists_userid.asp";
		var param = "Mobile=" + $("#Mobile_1").val() + $("#Mobile_2").val() + $("#Mobile_3").val();
		
		var IEIndex = navigator.appVersion.indexOf("MSIE");        // MSIE를 찾고 인덱스를 리턴
		var IE8Over = navigator.userAgent.indexOf("Trident");      // MS IE 8이상 버전 체크

		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){

					if( IEIndex > 0 || IE8Over > 0 )  {

					   var trident = navigator.userAgent.match(/Trident\/(\d.\d)/i);
					   if (trident != null){
						switch (trident[1]) {
						 case "7.0" :
						  strVer = "11.0";
							// 가입진행
							//alert('11');
							//doJoinSubmit2();
							$("#frmJoin").submit();
						  break;
						 case "6.0" :
						  strVer = "10.0";
							// 가입진행
							//alert('10');
							//doJoinSubmit2();
							$("#frmJoin").submit();
						  break;
						 case "5.0" :
						  strVer = "9.0";
							// 가입진행
							//alert('9');
							doJoinSubmit();
						  break;      
						 case "4.0" :
						  strVer = "8.0";
							// 가입진행
							//alert('8');
							doJoinSubmit();
						  break;
						 default :
							//alert('99');
						  break;
						}
					   }    
					   
					   //strNav = "Microsoft Internet Explorer " + strVer;    
					   
					  }  else  {   
					   strNav = "기타 브라우저";
					  }
						
					doJoinSubmit();

				}else{
					alert("이미 회원가입 되어있으시거나, 다른 사이트를 통해 HC고객센터에 지원하신 이력이 있습니다. \n\n회원가입 여부를 확인 바랍니다.");
					window.location.href = "/membership/find.asp"
				}
			}
			, error: function(xhr, status, error) {alert("[핸드폰번호 중복 체크] " + error);}
		});
	}
	

	function doChkMobilebtn() {

		var url = "/common/dev/ajax/check_exists_userid.asp";
		var param = "Mobile=" + $("#Mobile_1").val() +"-"+ $("#Mobile_2").val() +"-"+ $("#Mobile_3").val();


		if (!$("#frmJoin").find("#Mobile_1").val() || $("#frmJoin").find("#Mobile_1").val().length!=3 ){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_1").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_2").val() || $("#frmJoin").find("#Mobile_2").val().length<3){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_2").focus();
			return;
		} else if (!$("#frmJoin").find("#Mobile_3").val() || $("#frmJoin").find("#Mobile_3").val().length!=4){
			alert("핸드폰번호를 입력해주세요.");
			$("#frmJoin").find("#Mobile_3").focus();
			return;
		} else if (!IsMobile($("#frmJoin").find("#Mobile_1").val() + $("#frmJoin").find("#Mobile_2").val() + $("#frmJoin").find("#Mobile_3").val())){
			alert("핸드폰번호가 바르지 않습니다.");
			$("#frmJoin").find("#Mobile_1").focus();
			return;
		
		}
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){
					alert("등록되지 않은 휴대폰번호입니다. 회원가입 시도 바랍니다.");
				}else{
					alert("등록된 휴대폰번호입니다. 로그인 시도 바랍니다.");
				}
			}
			, error: function(xhr, status, error) {alert("[핸드폰번호 중복 체크] " + error);}
		});

	}

	function doJoinSubmit(){
		 //alert($("#CrBYear").val())

			$("#frmJoin").submit();
	}


	$(function(){
    //ajax form submit
    $('#frmJoin').ajaxForm({

            success: function(response,status){
                //성공후 서버에서 받은 데이터 처리
				
				if (response=="SUCC") {
					openLayer("/membership/signUpComplete.asp","", "w668");
				}else if (response=="EXISTS"){
					alert("[회원가입] 이미 등록된 휴대폰번호입니다.\n등록된 휴대폰번호로 로그인 시도 바랍니다.");
				}else{
					alert("[회원가입] 가입처리중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}

                //alert("업로드 성공!!");
            },
            error: function(error){
                //에러발생을 위한 code페이지
				 alert("[회원가입] " + error);
            }                               
        });
	});


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
		$("#Mobile_1").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_1").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_2").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_2").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#Mobile_3").keypress(function(e){ return OnlyNumber(e,false); });
		$("#Mobile_3").keyup(function(){ $(this).val( FillNumerics($(this).val())); });
		$("#btnChkMobile").click(doChkMobilebtn);
	});

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
	<div id="wrap" class="recruitView">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				
			<div class="inner"><div class="inner2">
				<div class="popCont easySignUp">

					<h2><img src="/images/pages/tit_signup.png" alt="간편회원가입" /></h2>
					<p class="note">
						<img src="/images/pages/txt_login_04_n.gif" alt="입력하신 핸드폰번호가 아이디로 적용됩니다." /><br />
						<img src="/images/pages/txt_login_05_n.gif" alt="개인정보는 채용정보 제공 시에만 활용됩니다." /><br />
						<img src="/images/pages/txt_login_06_n.gif" alt="간편회원가입 후 자동 로그인 되어집니다." />
					</p>
					
					<form id="frmJoin" name="frmJoin" method="post" action="/membership/easySignUp_proc.asp">
					<input type="hidden" id="Flag" name="Flag">
					<input type="hidden" id="HopePart" name="HopePart" value!="상담원">
					<input type="hidden" id="PageFlag" name="PageFlag" value="<%=PageFlag%>">
					<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
					
					<div class="board_write">
						<table>
							<caption>회원가입양식</caption>
							<colgroup>
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
							</colgroup>
							<tbody>
								<tr>
									<th><label for="UsrName"><span class="ess">* </span>이름</label></th>
									<td><input type="text" id="UsrName" name="UsrName" title="이름" value="" class="w110" maxlength="25" /> </td>
									<th><label for="Birthday_Y"><span class="ess">* </span>생년월일</label></th>
									<td>
										<select name="Birthday_Y" id="Birthday_Y" title="년" class="w55">
											<option value="">년</option>
											<% For yy = Year(Date()) To GLOBAL_SEL_YEAR Step -1  %>
											<option value="<%=yy%>"><%=yy%></option>
											<% Next %>
										</select>
										년
										<select name="Birthday_M" id="Birthday_M" title="월" class="w40">
											<option value="">월</option>
											<% For mm = 1 TO 12 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select>
										월
										<select name="Birthday_D" id="Birthday_D" title="일" class="w40">
											<option value="">일</option>
											<% For mm = 1 TO 31 %>
											<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
											<% Next %>
										</select>
										일
									</td>
								</tr>
								<tr>
									<th><label for="Mobile_1"><span class="ess">* </span>휴대폰번호</label></th>
									<td><input type="text" name="Mobile_1" id="Mobile_1" title="휴대폰번호 앞 3자리" class="w35" maxlength="3" style="ime-mode:disabled;" /> - 
										<input type="text" name="Mobile_2" id="Mobile_2" title="휴대폰번호 중간 4자리" class="w35" maxlength="4" style="ime-mode:disabled;" /> - 
										<input type="text" name="Mobile_3" id="Mobile_3" title="휴대폰번호 끝 4자리" class="w35" maxlength="4" style="ime-mode:disabled;" />&nbsp;<a href="javascript:;" id="btnChkMobile"><img src="/images/pages/btn_checkmembership_n.png" alt="회원가입여부 확인" /></a>
									</td>
									<th><label for="Gender"><span class="ess">* </span>성별</label></th>
									<td>
										<input type="radio" name="Gender" id="Gender1" value="1" checked >남
										<input type="radio" name="Gender" id="Gender2" value="2" >여
									</td>
								</tr>
								<tr>
									<th><label for="AddrSido"><span class="ess">* </span>주소</label></th>
									<td>
										<%
										Dim SidoSql : SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
										response.write " ": Call SB_MakeHTMLSelect(objDBCon, SidoSql, "AddrSido", "지역", "", "", "", " class=""w55"" ", False)
										%>
										<span id="divGugun">
										<select name="AddrGugun" id="AddrGugun" title="지역중분류" class="w85">
											<option value="">지역중분류</option>
										</select>
										</span>
									</td>
									<th><label for="Passwd"><span class="ess">* </span>비밀번호</label></th>
									<td><input type="password" name="Passwd" id="Passwd" title="비밀번호" class="w110" maxlength="30" /></td>
								</tr>
								<tr>
									<th><label for="Email">E-mail</label></th>
									<td colspan="3">
										
										<input type="hidden" id="Email" name="Email" title="E-mail 주소" value="" /> 
										<input type="text" id="email1" name="email1" class="w114" title="이메일 주소의 앳(@) 이전 정보 입력" maxlength="18" value=""  />
										@
										<input type="text" id="email2" name="email2" class="w144" title="이메일 주소의 앳(@) 이후 정보 직접입력" maxlength="30"  value="" />
										<select id='emailDomains' name='emailDomains' title="이메일 주소의 앳(@) 이후 정보 선택" style="width:160px;"  onchange="document.frmJoin.email2.value=this.value;document.frmJoin.Email.value=document.frmJoin.email1.value+'@'+this.value;">
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
									</td>
								</tr>
							</tbody>
						</table>
					</div>
					
					<div class="section">
					&nbsp;
					</div>

					<% If  Ubound(arrCdEzCr,2) >= 0 Then  %>
					<p class="note2">
						<img src="/images/pages/txt_login_07.gif" alt="경력이 있으신 곳에 체크 해주세요" />
					</p>
					<div class="check">
						<% For i = 0 To Ubound(arrCdEzCr,2) %>
						<p><input type="checkbox" name="EzCareer" id="EzCareer_<%=i+1%>" value="<%=FN_InputVal(arrCdEzCr(1,i))%>" /><label for="EzCareer_<%=i+1%>"><%=arrCdEzCr(1,i)%></label></p>
						<% Next %>
					</div>
					<% End If %>
					
					<div class="section">
					&nbsp;
					</div>

					<h3><img src="/images/pages/stit_resumeMod_02.jpg" alt="학력사항" /></h3>
					<div class="section">
						<div id="divEduInfo">
							<table class="tblForm alignC">
								<colgroup>
									<col class="w150" />
									<col />
								</colgroup>
								<tbody>
									
									<tr>
										<td>
											<select id="EduLast" name="EduLast" title="최종학력" class="w120">
												<option value="0">중학교</option>
												<option value="1">고등학교</option>
												<option value="2">대학(2,3년)</option>
												<option value="3">대학교(4년)</option>
												<option value="4">대학원</option>
											</select>
										</td>
										<td class="alignL">
											&nbsp;<input type="radio" name="EduType" id="EduType1" value="1" checked>졸업
											<input type="radio" name="EduType" id="EduType2" value="2" >재학중
											<input type="radio" name="EduType" id="EduType3" value="3" >휴학
											<input type="radio" name="EduType" id="EduType4" value="4" >중퇴
										</td>
									</tr>
								</tbody>
							</table>
						</div>
					</div>
					
					<div class="section">
					&nbsp;
					</div>

					<h3><img src="/images/pages/stit_resumeMod_03.jpg" alt="경력사항" /></h3>
						<div class="section">
							<div id="divCareerInfo">
								<table class="tblForm alignC">
									<colgroup>
										<col class="w300" />
										<col class="w100" />
										<col class="w100" />
										<col class="w100" />
										<col />
									</colgroup>
									 <thead>
										<tr>
											<th><label for="">근무기간</label></th>
											<th><label for="">직장명</label></th>
											<th><label for="">부서명</label></th>
											<th><label for="">담당업무</label></th>
											<th><label for="">퇴직사유</label></th>
										</tr>
									</thead>
									<tbody>
									<!-- <%For i=1 To 3 %>
										<tr>
											<td>
												<select id="CompanyYear_<%=i%>" name="CompanyYear_<%=i%>" title="근무개월수" class="w70">
													<option value="0">1년미만</option>
													<option value="1">1년이상</option>
													<option value="2">2년이상</option>
													<option value="3">3년이상</option>
													<option value="4">4년이상</option>
												</select>
											</td>
											<td class="alignC"><input type="text" id="CompanyName_<%=i%>" name="CompanyName_<%=i%>" title="직장명" value="" class="w125" maxlength="20" /></td>
											<td><input type="text" id="Task_<%=i%>" name="Task_<%=i%>" title="업무" value="" class="w80" maxlength="15" /></td>
										</tr>
									<%Next%> -->
									<%For i=1 To 3 %>
									<tr>
											<td>
												<select id="CrBYear_<%=i%>" name="CrBYear_<%=i%>" title="입사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrBMonth_<%=i%>" name="CrBMonth_<%=i%>" title="입사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
												~
												<select id="CrEYear_<%=i%>" name="CrEYear_<%=i%>" title="퇴사년도" class="w55">
													<% For yy = GLOBAL_SEL_YEAR TO Year(Date()) %>
													<option value="<%=yy%>"><%=yy%></option>
													<% Next %>
												</select>
												<select id="CrEMonth_<%=i%>" name="CrEMonth_<%=i%>" title="퇴사월" class="w40">
													<% For mm = 1 TO 12 %>
													<option value="<%=Right("0" & mm,2)%>"><%=Right("0" & mm,2)%></option>
													<% Next %>
												</select>
											</td>
											<td class="alignC"><input type="text" id="CompanyName_<%=i%>" name="CompanyName_<%=i%>" title="직장명" value="" class="w60" maxlength="20" /></td>
											<td class="alignC"><input type="text" id="Department_<%=i%>" name="Department_<%=i%>" title="부서" value="" class="w60" maxlength="15" /></td>
											<td><input type="text" id="Task_<%=i%>" name="Task_<%=i%>" name="Task_<%=i%>" title="업무" value="" class="w60" maxlength="15" /></td>
											<td><input type="text" id="Reason_<%=i%>" name="Reason_<%=i%>" name="Reason_<%=i%>" title="퇴사사유" value="" class="w60" maxlength="25" /></td>
										</tr>
									<%Next%>
									</tbody>
								</table>
							</div>
						</div>
					
					<div class="section">
					&nbsp;
					</div>
					
					<h3><img src="/images/pages/stit_resumeMod_06.jpg" alt="첨부파일" /></h3>
					<div class="section">
							<table class="board_write">
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
										<th><label for="ImgPath">사진(이미지)</label></th>
										<td colspan="5">
											<input type="file" id="ImgPath" name="ImgPath" title="이미지첨부" value="" class="w360" size="92" />
											<input type="hidden" id="DelImgPath" name="DelImgPath" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Attach_1">이력서 1</label></th>
										<td colspan="5">
											<input type="file" id="Attach_1" name="Attach_1" title="첨부파일 1" value="" class="w360" size="92" /> 
											<input type="hidden" id="DelAttach_1" name="DelAttach_1" value="">
										</td>
									</tr>
									<tr>
										<th><label for="Attach_2">이력서 2</label></th>
										<td colspan="5">
											<input type="file" id="Attach_2" name="Attach_2" title="첨부파일 2" value="" class="w360" size="92" /> 
											<input type="hidden" id="DelAttach_2" name="DelAttach_2" value="">
										</td>
									</tr>
								</tbody>
							</table>
					</div>

					<div class="section">
					&nbsp;
					</div>

					<h3><img src="/images/pages/stit_resumeMod_05.jpg" alt="자기소개서" /></h3>
					<div class="textAreas">
						<textarea name="Intro" id="Intro" rows="10" cols="40"></textarea>
					</div>


					<h3><img src="/images/pages/stit_resumeMng_01.jpg" alt="서비스 약관" /></h3>
					<div class="inner_txt">
						<h4>제1조 (목적)</h4>
						<p>본 약관은 (주)한국코퍼레이션에서 운영하는 취업정보 서비스에 대한 회원제도에 대한 제반사항을 규정하는데 목적이 있습니다.</p>
						<h4>제2조 (정의)</h4>
						<p>이 약관에서 사용되는 용어의 정의는 아래와 같습니다.</p>
						<ul>
							<li>1. 회원 : 회사와 서비스 이용에 관한 계약을 체결한 자</li>
							<li>2. 아이디(ID) : 회원 식별과 회원의 서비스 이용을 위하여 회원이 선정하고 회사가 승인하는 문자와 숫자의 조합</li>
							<li>3. 비밀번호 : 회원이 통신상의 자신의 비밀을 보호하기 위해 선정한 문자와 숫자의 조합</li>
							<li>4. 해지 : 회사 또는 회원이 서비스 이용 이후 그 이용계약을 종료시키는 의사표시</li>
							<li>5. 개인정보 : 개인에 관한 정보로서 해당 정보에 포함되어 있는 성명, 주민등록번호 등의 사항에 의하여 해당 개인을 식별할<br />&nbsp;&nbsp;&nbsp;&nbsp;수 있는 정보</li>
						</ul>
						<h4>제3조 (약관의 효력 및 변경)</h4>
						<ul>
							<li>① 이 약관은 회사 인터넷 사이트를 통하여 이를 공지하거나 기타의 방법으로 회원에게 통지함으로써 효력이 발생됩니다. </li>
							<li>② 회사는 사정상 중요한 사유가 발생될 경우 사전 고지 없이 이 약관의 내용을 변경할 수 있으며, 변경된 약관은 제1항과 같은<br />&nbsp;&nbsp;&nbsp;&nbsp;방법으로 공지 또는 통지함으로써 효력이 발생됩니다. </li>
							<li>③ 회원은 변경된 약관에 동의하지 않을 경우 회원 탈퇴를 요청할 수 있으며, 변경된 약관의 효력 발생일 이후에도 서비스를<br />&nbsp;&nbsp;&nbsp;&nbsp;계속 사용할 경우 약관의 변경 사항에 동의한 것으로 간주됩니다.</li>
						</ul>
						<h4>제4조 (약관 외 준칙)</h4>
						<p>본 약관에 명시되지 아니한 사항에 대해서는 전기통신기본법, 전기통신사업법, 정보통신망 이용촉진등에 관한 법률, 전자거래기본법, 신용정보이용및보호에관한법률, 기타 관련 법령의 규정에 따릅니다.</p>
						<h4>제5조 (이용 계약의 성립)</h4>
						<ul>
							<li>① 회사의 회원으로 가입하여 회사가 제공하는 서비스를 받고자 하는 자는 회원가입절차를 거쳐 회원으로 가입하여야 하며,<br />&nbsp;&nbsp;&nbsp;&nbsp;회원 가입희망자가 이 약관의 내용을 숙지하고 '동의' 버튼을 누르면 이 약관에 동의하는 것으로 간주됩니다.<br />&nbsp;&nbsp;&nbsp;&nbsp;약관 변경 시에도 이와 동일하며, 변경된 약관에 동의하지 않을 경우 회원 등록 취소가 가능합니다. </li>
							<li>② 만 14세 미만의 이용자가 회원으로 가입하여 서비스를 이용하고자 하는 경우에는 부모 등 법정 대리인의 동의를 얻은<br />&nbsp;&nbsp;&nbsp;&nbsp;후에 서비스 이용을 신청을 하여야 합니다.</li>
							<li>③ 이용계약은 서비스 이용 희망자의 이용약관 동의 후 이용 신청에 대하여 회사가 승낙함으로써 성립합니다.</li>
						</ul>
						<h4>제6조 (이용신청)</h4>
						<ul>
							<li>① 회원으로 가입하여 서비스를 이용하기를 희망하는 자는 회사가 요청하는 소정의 가입신청 양식에서 요구하는 사항을<br />&nbsp;&nbsp;&nbsp;&nbsp;기록하여 신청합니다.</li>
							<li>② 온라인 가입신청 양식에 기재하는 모든 회원 정보는 실제 데이터인 것으로 간주하므로 실명이나 실제 정보를 입력하지<br />&nbsp;&nbsp;&nbsp;&nbsp;않은 사용자는 법적인 보호를 받을 수 없으며, 서비스 사용의 제한을 받으실 수도 있습니다.</li>
						</ul>
						<h4>제7조 (개인정보의 보호)</h4>
						<p>회사는 인터넷에서의 개인정보 보호를 매우 중요하게 생각하고 회원이 서비스를 이용함에 있어서 온라인 상에서 회사에게 제공한 개인정보가 보호받을 수 있도록 최선을 다하며, 전기통신기본법, 전기통신사업법, 정보통신망 이용 촉진 등에 관한 법률, 전자거래기본법, 신용정보이용및보호에관한법률, 기타 관련 법령에서 정하고 있는 정보통신 서비스 제공자의 개인정보 보호의무를 준수합니다.개인정정보호와 관련된 보다 상세한 내용은 별도의 회사 개인정보보호정책으로 규율 합니다. </p>
						<ul>
							<li>① 회원의 개인정보에 대해서는 회사의 개인정보 보호정책이 적용됩니다.</li>
							<li>② 당사의 회원 정보는 다음과 같이 수집, 사용, 관리, 보호됩니다.
								<ul>
									<li>1. 개인정보의 수집 : 회사는 귀하의 회사 서비스 가입 시 귀하가 제공하는 개인정보, 이력서 정보 등을 수집합니다. 
									<li>2. 개인정보의 사용 : 회사는 서비스 제공과 관련해서 수집된 회원의 개인정보를 본인의 승낙 없이 제3자에게 누설,<br />&nbsp;&nbsp;&nbsp;&nbsp;배포하지 않습니다. </li>
									<li>&nbsp;&nbsp;&nbsp;&nbsp;단, 아래 경우에는 그러하지 않습니다.
										<ul>
											<li>1) 전기통신기본법 등 법률의 규정에 의해 국가기관의 요구가 있는 경우</li>
											<li>2) 범죄에 대한 수사상의 목적이 있거나 정보통신윤리 위원회의 요청이 있는 경우 또는 기타 관계법령에서 정한<br />&nbsp;&nbsp;&nbsp;&nbsp;절차에 따른 요청이 있는 경우</li>
											<li>3) 구직회원에 가입한 미성년자의 법정대리인이 정보를 요청하여 서류상으로 가족관계가 확인된 경우</li>
											<li>4) 회원이 당사에 제공한 개인정보를 스스로 공개한 경우</li>
										</ul>
									</li>
									<li>3. 개인정보의 보호 : 회원의 주소, E-Mail주소, 휴대폰번호, 직장명, 소속부서, 직위, 우편물 수령장소 등의 회원 개인정보는<br />&nbsp;&nbsp;&nbsp;&nbsp;회사의 책임으로 관리·이용합니다.</li>
								</ul>
								<li>③ 회사는 업무와 관련하여 회원 전체, 일부의 개인 정보에 관한 자료를 작성하여 이를 사용할 수 있고, 서비스를 통하여<br />&nbsp;&nbsp;&nbsp;&nbsp;회원의 컴퓨터에 쿠키를 전송할 수 있습니다.</li>
							</li>
						</ul>
						<h4>제8조 (이용 신청의 승낙)</h4>
						<ul>
							<li>① 회사는 제6조에 따른 이용자의 이용신청에 대하여 신용정보이용 및 보호에 관한 법률이 정하는 신용불량자로 등록되어<br />&nbsp;&nbsp;&nbsp;&nbsp;있는 등의 특별한 사정이 없는 한 접수 순서대로 이용 신청을 승낙합니다.</li>
							<li>② 회사는 다음 각 호의 1에 해당하는 경우 이용신청에 대한 승낙을 제한할 수 있고, 그 사유가 해소될 때까지 승낙을 유보할<br />&nbsp;&nbsp;&nbsp;&nbsp;수 있습니다.
								<ul>
									<li>1. 서비스 관련 설비에 여유가 없는 경우 </li>
									<li>2. 기술상 지장이 있는 경우</li>
									<li>3. 기타 회사의 사정상 필요하다고 인정되는 경우</li>
								</ul>
							</li>
							<li>③ 회사는 다음 각 호의 1에 해당하는 이용계약 신청에 대하여는 이를 승낙하지 아니 할 수 있습니다.
								<ul>
									<li>1. 본인의 실명으로 신청하지 않은 경우</li>
									<li>2. 다른 사람의 명의를 사용하여 신청한 경우</li>
									<li>3. 이용 신청 시 필요 사항을 허위로 기재하여 신청한 경우</li>
									<li>4. 신용정보이용및보호에관한법률이 정하는 신용불량자로 등록되어 있는 경우</li>
									<li>5. 사회의 안녕과 질서 혹은 미풍양속을 저해할 목적으로 신청한 경우</li>
									<li>6. 이 약관 제22조 제 2 항에 따라 회원자격을 상실한 적이 있는 경우</li>
									<li>7. 기타 회사가 정한 이용 신청 요건이 미비 된 경우</li>
								</ul>
							</li>
							<li>④ 제2항 또는 제3항에 의하여 이용신청의 승낙을 유보하거나 승낙하지 아니하는 경우, 회사는 이를 이용신청자에게<br />&nbsp;&nbsp;&nbsp;&nbsp;알려야 합니다. 다만, 회사 사정상 이용신청자에게 통지할 수 없는 경우는 예외로 합니다.</li>
						</ul>
						<h4>제9조 (계약 사항의 변경)</h4>
						<ul>
							<li>① 회원은 개인정보관리를 통해 언제든지 본인의 개인정보를 열람하고 수정할 수 있습니다.</li>
							<li>② 회원은 이용신청 시 기재한 사항이 변경되었을 경우 온라인으로 수정을 해야 하며 회원정보의 미변경으로 인하여 발생되는<br />&nbsp;&nbsp;&nbsp;&nbsp;문제의 책임은 회원에게 있습니다.</li>
						</ul>
						<h4>제 10 조 (회사의 의무)</h4>
						<ul>
							<li>① 회사는 특별한 사정이 없는 한 회원이 서비스 이용 신청 후, 48시간 이내에 서비스를 이용할 수 있도록 합니다.</li>
							<li>② 회사는 이 약관에서 정한 바에 따라 계속적이고 안정적인 서비스의 제공을 위하여 지속적으로 노력하며, 설비에 장애가<br />&nbsp;&nbsp;&nbsp;&nbsp;생기거나 멸실 된 때에는 지체 없이 이를 수리 복구하여야 합니다.<br />&nbsp;&nbsp;&nbsp;&nbsp;다만, 천재지변, 비상사태 또는 그밖에 부득이한 경우에는 그 서비스를 일시 중단하거나 중지할 수 있습니다.</li>
							<li>③ 회사는 회원으로부터 소정의 절차에 의해 제기되는 의견이나 불만이 정당하다고 인정할 경우에는 적절한 절차를<br />&nbsp;&nbsp;&nbsp;&nbsp;거처 처리하여야 합니다. 처리 시 일정 기간이 소요될 경우 회원에게 그 사유와 처리 일정을 알려주어야 합니다.</li>
							<li>④ 회사는 이용계약의 체결, 계약사항의 변경 및 해지 등 이용고객과의 계약 관련 절차 및 내용 등에 있어 이용고객에게 편의를<br />&nbsp;&nbsp;&nbsp;&nbsp;제공하도록 노력합니다.</li>
						</ul>
						<h4>제11조 (회원의 의무)</h4>
						<ul>
							<li>① 회원은 이 약관에서 규정하는 사항과 서비스 이용안내 또는 주의사항 등 회사가 공지 혹은 통지하는 사항을 준수하여야 하며,<br />&nbsp;&nbsp;&nbsp;&nbsp;기타 회사의 업무에 방해되는 행위를 하여서는 아니 됩니다.</li>
							<li>② 회원의 ID와 비밀번호에 관한 모든 관리책임은 회원에게 있습니다. 회원에게 부여된 ID와 비밀번호의 관리 소흘,<br />&nbsp;&nbsp;&nbsp;&nbsp;부정사용에 의하여 발생하는 모든 결과에 대한 책임은 회원에게 있습니다.</li>
							<li>③ 회원은 자신의 ID나 비밀번호가 부정하게 사용되었다는 사실을 발견한 경우에는 즉시 회사에 신고하여야 하며, 신고를<br />&nbsp;&nbsp;&nbsp;&nbsp;하지 않아 발생하는 모든 결과에 대한 책임은 회원에게 있습니다.</li>
							<li>④ 회원은 내용별로 회사가 서비스 공지사항에 게시하거나 별도로 공지한 이용제한 사항을 준수하여야 합니다.</li>
							<li>⑤ 회원은 회사의 사전승낙 없이는 서비스를 이용하여 영업활동을 할 수 없으며, 그 영업활동의 결과와 회원이 약관에 위반한<br />&nbsp;&nbsp;&nbsp;&nbsp;영업활동을 하여 발생한 결과에 대하여 회사는 책임을 지지 않습니다.<br />&nbsp;&nbsp;&nbsp;&nbsp;회원은 이와 같은 영업활동으로 회사가 손해를 입은 경우 회원은 회사에 대하여 손해배상의무를 집니다.</li>
							<li>⑥ 회원은 회사의 명시적인 동의가 없는 한 서비스의 이용권한, 기타 이용계약상 지위를 타인에게 양도, 증여할 수 없으며,<br />&nbsp;&nbsp;&nbsp;&nbsp;이를 담보로 제공할 수 없습니다.</li>
							<li>⑦ 회원은 서비스 이용과 관련하여 다음 각 호의 1에 해당되는 행위를 하여서는 아니 됩니다.
								<ul>
									<li>1. 다른 회원의 ID와 비밀번호, 주민등록번호 등을 도용하는 행위 
									<li>2. 본 서비스를 통하여 얻은 정보를 회사의 사전승낙 없이 회원의 이용 이외 목적으로 복제하거나 이를 출판 및 방송 등에<br />&nbsp;&nbsp;&nbsp;&nbsp;사용하거나 제3자에게 제공하는 행위 </li>
									<li>3. 타인의 특허, 상표, 영업비밀, 저작권 기타 지적재산권을 침해하는 내용을 게시, 전자메일 또는 기타 방법으로 타인에게<br />&nbsp;&nbsp;&nbsp;&nbsp;유포하는 행위 </li>
									<li>4. 공공질서 및 미풍양속에 위반되는 저속, 음란한 내용의 정보, 문장, 도형 등을 전송, 게시, 전자메일 또는 기타의 방법으로<br />&nbsp;&nbsp;&nbsp;&nbsp;타인에게 유포하는 행위 </li>
									<li>5. 모욕적이거나 위협적이어서 타인의 프라이버시를 침해할 수 있는 내용을 전송, 게시, 전자메일 또는 기타의 방법으로<br />&nbsp;&nbsp;&nbsp;&nbsp;타인에게 유포하는 행위 </li>
									<li>6. 해킹, 바이러스를 유포하거나 타인의 의사에 반해 광고성 정보 등 일정한 내용을 지속적으로 전송하는 행위 </li>
									<li>7. 회사의 직원이나 관리자를 가장하거나 사칭하여 내용물을 게시, 등록하거나 메일을 발송하는 행위</li>
									<li>8. 범죄와 결부된다고 객관적으로 판단되는 행위</li>
									<li>9. 회사의 승인을 받지 않고 다른 사용자의 개인정보를 수집 또는 저장하는 행위 </li>
									<li>10. 기타 관계법령에 위배되는 행위</li>
								</ul>
							</li>
						</ul>
						<h4>제12조 (정보의 제공)</h4>
						<p>회사는 회원이 서비스 이용 중 필요가 있다고 인정되는 다양한 정보를 공지사항이나 전자우편 등의 방법으로 회원에게 제공할 수 있습니다.</p>
						<h4>제13조 (회원의 게시물)</h4>
						<p>회사는 회원이 게시하거나 등록하는 서비스내의 내용물이 다음 각 호의 1에 해당한다고 판단되는 경우에 사전통지 없이 삭제할 수 있습니다.</p>
						<ul>
							<li>1. 다른 회원 또는 제3자를 비방하거나 중상모략으로 명예를 손상시키는 내용인 경우</li>
							<li>2. 공공질서 및 미풍양속에 위반되는 내용인 경우 </li>
							<li>3. 범죄적 행위에 결부된다고 인정되는 내용일 경우 </li>
							<li>4. 회사의 저작권, 제3자의 저작권 등 기타 권리를 침해하는 내용인 경우 </li>
							<li>5. 회사가 규정한 게시기간을 초과한 경우 </li>
							<li>6. 게시판의 성격에 부합하지 않는 게시물의 경우</li>
							<li>7. 기타 관계법령에 위반된다고 판단되는 경우</li>
						</ul>
						<h4>제14조 (게시물의 저작권)</h4>
						<p>서비스에 게재된 자료에 대한 권리는 다음과 같습니다.</p>
						<ul>
							<li>1. 게시물에 대한 권리와 책임은 게시자에게 있으며 회사는 게시자의 동의 없이는 이를 서비스 내 게재 이외에 영리적 목적으로<br />&nbsp;&nbsp;&nbsp;&nbsp;사용할 수 없습니다. 단, 비영리적인 경우에는 그러하지 아니하며 또한 회사는 서비스내의 게재권을 갖습니다.</li>
							<li>2. 회원은 서비스를 이용하여 얻은 정보를 가공, 판매하는 행위 등 서비스에 게재된 자료를 상업적으로 사용할 수 없습니다.</li>
						</ul>
						<h4>제 15 조 (서비스의 내용)</h4>
						<ul>
							<li>① 회사가 제공하는 서비스의 내용은 아래와 같습니다.
								<ul>
									<li>1. 회원용 취업정보 서비스</li>
									<li>2. 전자우편(E-mail) 서비스</li>
									<li>3. SMS 문자 서비스 </li>
									<li>4. 온라인 이력서 접수서비스 등</li>
								</ul>
							</li>
							<li>② 회사는 필요한 경우 서비스의 내용을 추가 또는 변경할 수 있습니다. 이 경우 회사는 추가 또는 변경내용을 웹사이트를<br />&nbsp;&nbsp;&nbsp;&nbsp;통해 공지합니다.</li>
						</ul>
						<h4>제16조(서비스의 요금)</h4>
						<p>회사가 개인에게 제공하는 서비스는 무료를 원칙으로 합니다. 다만 회사는 추후에 일부 서비스를 유료로 할 수 있으며, 그 자세한 내용에 대하여는 웹사이트를 공지합니다.</p>
						<h4>제17조(서비스의 개시)</h4>
						<p>서비스는 회사가 제8조에 따라서 이용신청을 승낙한 때로부터 즉시 개시됩니다.다만, 회사의 업무상 또는 기술상의 장애로 인하여 서비스를 즉시 개시하지 못하는 경우 회사는 회원에게 이를 통지하거나 당 사이트를 통해 공지합니다.</p>
						<h4>제18조 (서비스 이용시간)</h4>
						<ul>
							<li>① 서비스의 이용은 회사의 업무상 또는 기술상 특별한 지장이 없는 한 연중무휴 1일 24시간 가능함을 원칙으로 합니다.<br />&nbsp;&nbsp;&nbsp;&nbsp;다만 정기 점검 등의 필요로 회사가 정한 날이나 시간은 그러하지 않습니다.</li>
							<li>② 회사는 서비스를 일정범위로 분할하여 각 범위별로 이용가능 시간을 별도로 정할 수 있습니다.<br />&nbsp;&nbsp;&nbsp;&nbsp;이 경우 사전에 공지를	통해 그 내용을 알립니다.</li>
						</ul>
						<h4>제19조 (서비스 이용 책임)</h4>
						<p>회원은 회사에서 권한 있는 사원이 서명한 명시적인 서면에 구체적으로 허용한 경우를 제외하고는 서비스를 이용하여 상품을 판매하는 영업활동을 할 수 없으며 특히 해킹, 돈벌이 광고, 음란 사이트 등을 통한 상업행위, 상용S/W 불법배포 등을 할 수 없습니다. 이를 어기고 발생한 영업활동의 결과 및 손실, 관계기관에 의한 구속 등 법적 조치 등에 관해서는 회사가 책임을 지지 않습니다.</p>
						<h4>제20조 (서비스 제공의 중지 등)</h4>
						<ul>
							<li>① 회사는 다음에 해당하는 경우 서비스 제공을 중지할 수 있습니다.
								<ul>
									<li>1. 서비스용 설비의 보수 등 공사로 인한 부득이한 경우</li>
									<li>2. 전기통신사업법에 규정된 기간통신사업자가 전기통신 서비스를 중지했을 경우 </li>
									<li>3. 기타 불가항력적 사유가 있는 경우</li>
								</ul>
							</li>
							<li>② 회사는 국가비상사태, 정전, 서비스 설비의 장애 또는 서비스 이용의 폭주 등으로 정상적인 서비스 이용에 지장이 있는<br />&nbsp;&nbsp;&nbsp;&nbsp;때에는 서비스의 전부 또는 일부를 제한하거나 중지할 수 있습니다.</li>
							<li>③ 회사는 제1항 및 제2항의 규정에 의하여 서비스의 이용을 제한하거나 중지한 때에는 그 사유 및 제한기간 등을 회사의<br />&nbsp;&nbsp;&nbsp;&nbsp;사이트를 통하여 공지하거나 기타의 방법으로 지체 없이 회원에게 알려야 합니다.</li>
						</ul>
						<h4>제21조 (계약 해지 및 이용 제한)</h4>
						<ul>
							<li>① 회원이 이용 계약을 해지하고자 하는 경우에는 회원 본인이 온라인을 통해 회원탈퇴를 회사에 신청하여야 합니다. 
							<li>② 회사는 회원이 다음에 해당하는 행위를 하였을 경우 사전통지 없이 이용계약을 해지하거나 또는 기간을 정하여 서비스<br />&nbsp;&nbsp;&nbsp;&nbsp;이용을 중지할 수 있습니다.
								<ul>
									<li>1. 타인의 개인정보, ID 및 비밀번호를 도용한 경우</li>
									<li>2. 가입한 이름이 실명이 아닌 경우</li>
									<li>3. 같은 사용자가 다른 ID로 이중등록을 한 경우</li>
									<li>4. 타인의 명예를 손상시키거나 불이익을 주는 행위를 한 경우</li>
									<li>5. 회사, 다른 회원 또는 제 3자의 지적재산권을 침해하는 경우</li>
									<li>6. 공공질서 및 미풍양속에 저해되는 내용을 고의로 유포시킨 경우</li>
									<li>7. 회원이 국익 또는 사회적 공익을 저해할 목적으로 서비스 이용을 계획 또는 실행하는 경우</li>
									<li>8. 서비스 운영을 고의로 방해한 경우</li>
									<li>9. 서비스의 안정적 운영을 방해할 목적으로 다량의 정보를 전송하거나 광고성 정보를 전송하는 경우</li>
									<li>10. 정보통신설비의 오작동이나 정보의 파괴를 유발시키는 컴퓨터 바이러스 프로그램 등을 유포하는 경우</li>
									<li>11. 정보통신윤리위원회 등 외부기관의 시정요구가 있거나 불법선거운동과 관련하여 선거관리위원회의 유권해석을<br />&nbsp;&nbsp;&nbsp;&nbsp;받은 경우</li>
									<li>12. 회사의 서비스 정보를 이용하여 얻은 정보를 회사의 사전 승낙 없이 복제 또는 유통시키거나 상업적으로 이용하는 경우</li>
									<li>13. 회원이 자신의 홈페이지와 게시판에 음란물을 게재하거나 음란 사이트 링크하는 경우</li>
									<li>14. 본 약관을 포함하여 기타 회사가 정한 이용 조건에 위반한 경우</li>
								</ul>
							</li>
						</ul>
						<h4>제22조 (손해배상 및 면책)</h4>
						<ul>
							<li>① 회사는 천재지변 또는 이에 준하는 불가항력으로 인하여 서비스를 제공할 수 없는 경우에는 서비스 제공에 관한<br />&nbsp;&nbsp;&nbsp;&nbsp;책임이 면제됩니다.</li>
							<li>② 회사는 회원의 귀책사유로 인한 서비스 이용의 장애에 대하여 책임을 지지 않습니다.</li>
							<li>③ 회사는 회원이 서비스를 이용하여 기대하는 수익을 상실한 것이나 서비스를 통하여 얻은 자료로 인한 손해에 관하여<br />&nbsp;&nbsp;&nbsp;&nbsp;책임을 지지 않습니다.</li>
							<li>④ 회사는 회원이 서비스에 게재한 정보, 자료, 사실의 신뢰도, 정확성 등 내용에 관하여는 책임을 지지 않습니다.</li>
							<li>⑤ 회사는 서비스 이용과 관련하여 가입자에게 발생한 손해 가운데 가입자의 고의, 과실에 의한 손해에 대하여 책임을<br />&nbsp;&nbsp;&nbsp;&nbsp;지지 않습니다.</li>
						</ul>
						<h4>제23조 (관할법원)</h4>
						<p>서비스 이용으로 발생한 분쟁에 대해 소송이 제기될 경우 서울지방법원을 전속 관할법원으로 합니다.</p>
						<h4>부 칙</h4>
						<p>(시행일) 이 약관은 2006년 2월 1일부터 시행합니다.</p>
					</div>
					<p class="chk"><input type="checkbox" id="chk_consent1" name="chk_consent1" title="약관동의" /><label for="chk_consent1">위의 서비스약관에 동의합니다.</label></p>
					

					<div class="section">
					&nbsp;
					</div>

					<h3><img src="/images/pages/stit_resumeMng_02.jpg" alt="개인정보 수집동의" /></h3>
					<div class="inner_txt">
						
						
						<p><strong>아래의 개인정보 취급방침을 반드시 읽고 회원가입 절차를 이용해주시기 바랍니다.</strong></p>
						
						
						<h4>[개인정보 수집 및 이용 안내]</h4>
						(주)한국코퍼레이션는 개인정보보호법, 정보통신망 이용 촉진 및 정보보호 등에 관한 법률 등 관련 법령상의 개인정보보호 규정을 준수하며, 지원자의 개인정보보호에 최선을 다하고 있습니다.<br />
						한국코퍼레이션 사이트 내에서 주민등록번호를 취득하거나 이용하지 않습니다.<br />
						한국코퍼레이션의 기가입자의 주민등록번호는 모두 폐기합니다.<br />
						한국코퍼레이션은 개인식별을 위하여 이름/생년월일/휴대폰전화번호를 주민등록번호의 대체수단으로 이용합니다.<br />
						한국코퍼레이션 사이트는 본인의 동의 없이 개인정보를 수집 및 활용하거나 제3자에게 제공하지 않습니다.<br />
						<h4>제1조 개인정보의 수집 및 이용목적</h4>
							(주)한국코퍼레이션에서 운영하는 고객센터의 채용 지원과 자격요건의 확인 및 원활한 채용 진행, 진행단계별 결과, 채용관련 정보를 안내하기 위함.

						<h4>제2조 수집하는 개인정보의 항목</h4>
							1) 필수 : 휴대폰전화번호(=로그인 ID), 이름, 생년월일, 성별, 주소<br />
							2) 선택 : 사진 등록, E-mail주소, 학력사항, 경력사항, 자기소개서 등

						<h4>제3조 개인정보의 보유 및 이용기간</h4>
						   지원자의 개인정보는 원칙적으로 개인정보의 수집 및 이용목적이 달성되면 지체 없이 파기합니다<br />
						   다만, 다음의 정보에 대해서는 아래의 이유로 명기한 기간 동안 보유합니다.<br />
						   가. 회사 내부 방침에 의한 정보보유<br />
						   □ 보유사유 : 채용관련 정보 제공<br />
						   □ 보유정보 : 휴대폰전화번호(=로그인 ID), 이름, 생년월일, 성별, 주소<br />
							사진, E-mail주소, 학력사항, 경력사항, 자기소개서 등 입사지원서<br />
						   □ 보유기간 : 입사지원일로부터 3년<br />

						<h4>제4조 동의를 거부할 권리 및 동의 거부에 따른 제한사항</h4>

						   지원자는 위의 내용에 대해서 동의를 하지 않을 권리가 있습니다. 다만, 지원서를 통해 제공 받는 개인정보 제공에 동의하지 않을 경우 채용 전형이 제한될 수 있음을 알려드립니다. 
					</div>
					<p class="chk"><input type="checkbox" id="chk_consent2" name="chk_consent2" title="약관동의" /><label for="chk_consent2">위의 개인정보 수집에 동의합니다.</label></p>

					
					
					</form>
					<p class="popBtn">
						<a href="javascript:chkJoinForm();"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
						<a href="#" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
					</p>
				</div>
			</div>
		</div>

			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>