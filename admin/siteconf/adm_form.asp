<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<!-- #include virtual="/common/Seed/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : adm_form.asp / 관리자 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag
	Dim AdmID, AdmName, AdmLevel, AdmTel, AdmMobile, AdmEmail

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		
		If Session("ALevel") = "3" And CStr(Session("ASeq")) <> CStr(Seq) Then
			Call SB_ReturnErr("[권한] 잘못된 경로로 접근하셨습니다.","CLOSE")
		End If
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspAdm_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@AdmSeq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				AdmID 		= FN_InputVal(Rs("AdmID"))
				AdmName 	= FN_InputVal(Rs("AdmName"))
				AdmLevel 	= Rs("AdmLevel")
				AdmTel 		= FN_InputVal(Rs("AdmTel"))
				AdmMobile 	= FN_InputVal(Rs("AdmMobile"))
				AdmEmail 	= FN_InputVal(Rs("AdmEmail"))
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>관리자 추가 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			<% If Session("ALevel") = "1" Then %>
			if (!$("#AdmLevel").val()){
				alert("관리자 구분을 선택해주세요.");
				$("#AdmLevel").focus();
				return;
			}
			<% End If %>
			if (!$("#AdmID").val()){
				alert("아이디를 입력해주세요.");
				$("#AdmID").focus();
				return;
			}
			<% If Flag = "ADD" Then %>
			if (!$("#AdmPwd").val()){
				alert("비밀번호를 입력해주세요.");
				$("#AdmPwd").focus();
				return;
			}

			if(!fn_pw_check()){
				return false;
			}

			<% End If %>

			<% If Flag = "MOD" Then %>

			if ($("#AdmPwd").val()){
				if(!fn_pw_check()){
				return false;
				}
			}
			<% End If %>
			if (!$("#AdmName").val()){
				alert("이름을 입력해주세요.");
				$("#AdmName").focus();
				return;
			}
		
			$("#frmInfo").attr({action:"https://www.mpcjob.co.kr/admin/siteconf/adm_proc.asp", method:"post"}).submit();
		}
		
		$(function() {
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); })
			$("#btnSubmit").click(chkForm);
			$("#btnPassReset").click(function(){ $("#Flag").val('PWD'); $("#frmInfo").attr({action:"https://www.mpcjob.co.kr/admin/siteconf/adm_proc.asp", method:"post"}).submit(); });
		});
	//-->
	</script>

	<script type="text/javascript">

     var pw_passed = true;  // 추후 벨리데이션 처리시에 해당 인자값 확인을 위해

     function pwInputValueCheck(){  //특정문자 체크
    	  var InputValue = ["password", "admin", " ", ",", ";", ":", "'", "\\" , "\"" , "|"];
    	  var validpw = $("#AdmPwd").val(); 
    	  for(i=0; i<InputValue.length; i++){
    	   if(validpw.indexOf(InputValue[i]) >= 0){
    	    return InputValue[i];
    	   }
    	  }
    	  return null;
    }


    function fn_pw_check() {
        
        var pw = $("#AdmPwd").val();

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

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<h2>관리자 등록 (관리자를 추가하거나 수정하실 수 있습니다.)</h2>
			<div class="tblArea">
				<form id="frmInfo" name="frmInfo" action="adm_action.asp" method="post">
				<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
				<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
				<table class="tblForm">
					<colgroup>
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
						</tr>
						<% If Session("ALevel") = "1" Then %>
						<tr>
							<th><label for="">관리자 구분</label></th>
							<td colspan="3">
								<select id="AdmLevel" name="AdmLevel" title="관리자 구분" class="w340">
									<option value="">선택</option>
									<option value="1"<%=Fn_SetDefault(AdmLevel,"1"," selected","")%>>최고관리자</option>
									<option value="2"<%=Fn_SetDefault(AdmLevel,"2"," selected","")%>>최고관리자2</option>
									<option value="3"<%=Fn_SetDefault(AdmLevel,"3"," selected","")%>>일반관리자</option>
								</select> 
							</td>
						</tr>
						<% End If %>
						<tr>
							<th><label for="AdmID">아이디</label></th>
							<td>
								<input type="text" id="AdmID" name="AdmID" title="아이디" maxlength="30" value="<%=AdmID%>" class="w335" style="ime-mode:disabled;" <%=Fn_SetDefault(Flag,"MOD","readonly","")%> /> 
							</td>
							<th><label for="AdmPwd">비밀번호</label></th>
							<td>
								<input type="password" id="AdmPwd" name="AdmPwd" title="비밀번호" maxlength="30" value="" class="w335" /> 
							</td>
						</tr>
						<tr>
							<th><label for="AdmName">이름</label></th>
							<td>
								<input type="text" id="AdmName" name="AdmName" title="이름" maxlength="15" value="<%=AdmName%>" class="w335" /> 
							</td>
							<th><label for="AdmEmail">이메일</label></th>
							<td>
								<input type="text" id="AdmEmail" name="AdmEmail" title="이메일" maxlength="50" value="<%=AdmEmail%>" class="w335" style="ime-mode:disabled;" /> 
							</td>
						</tr>
						<tr>
							<th><label for="AdmTel">연락처</label></th>
							<td>
								<input type="text" id="AdmTel" name="AdmTel" title="연락처" value="<%If AdmTel <> "" Then Response.write oSeed.Decrypt(AdmTel) End If %>" class="w335" style="ime-mode:disabled;" /> 
							</td>
							<th><label for="AdmMobile">휴대폰</label></th>
							<td>
								<input type="text" id="AdmMobile" name="AdmMobile" title="휴대폰" value="<%If AdmMobile <> "" Then Response.write oSeed.Decrypt(AdmMobile) End If %>" class="w335" style="ime-mode:disabled;" /> 
							</td>
						</tr>
					</tbody>
				</table>
				</form>
			</div>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
					</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
					</span><span class="btns btn_b"><a href="#btnPassReset" id="btnPassReset">비밀번호횟수초기화</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">x 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>