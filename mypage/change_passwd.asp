<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : change_passwd.asp / HC - 비밀번호 변경 (Layer Popup)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-26 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<script type="text/javascript">
	function chkModForm(){
		
		if (!$("#frmModify").find("#Passwd").val()){
			alert("현재 비밀번호를 입력해주세요.");
			$("#frmModify").find("#Passwd").focus();
			return;
		} else if (!$("#frmModify").find("#New_Passwd").val()){
			alert("새 비밀번호를 입력해주세요.");
			$("#frmModify").find("#New_Passwd").focus();
			return;
		}else if(!fn_pw_check()){
				return false;
		}else if (!$("#frmModify").find("#Re_Passwd").val()){
			alert("새 비밀번호를 다시 입력해주세요.");
			$("#frmModify").find("#Re_Passwd").focus();
			return;
		} else if ($("#frmModify").find("#New_Passwd").val() != $("#frmModify").find("#Re_Passwd").val()){
			alert("입력하신 새 비밀번호와 새 비밀번호 확인이 동일하지 않습니다.");
			$("#frmModify").find("#New_Passwd").focus();
			return;
		}
		
		doChkPasswd();
	}
		
	function doChkPasswd() {
		var url = "/common/dev/ajax/check_user_passwd.asp";
		var param = "Passwd=" + $("#frmModify").find("#Passwd").val();
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				if (result=="true"){
					// 변경진행
					doModSubmit();
				}else{
					// 비밀번호 불일치
					alert("입력하신 비밀번호가 맞지 않습니다.");
				}
			}
			, error: function(xhr, status, error) {alert("[비밀번호 체크] " + error);}
		});
	}

	function doModSubmit(){
		var url = "change_proc.asp"
		param = $("#frmModify").serialize();
		
		$.ajax({
			type: "POST"
			,crossDomain:true
		    ,dataType:"jsonp"
		    ,jsonpCallback:"callback"
			, url: url
			, data: param
			, success: function(content){ 
				if (content.result=="SUCC") {
					alert("비밀번호 변경이 완료 되었습니다.");
					window.location.reload();
				}else{
					alert("[비밀번호 변경] 비밀번호 변경중 장애가 발생했습니다. 관리자에게 문의 바랍니다.");
				}
			}
			, error: function(xhr, status, error) { alert("[비밀번호] " + error);}
		});
	}
	
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
        
        var pw = $("#New_Passwd").val();

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
	<div class="inner"><div class="inner2">
		<div class="popCont easySignUp">
			<h2><img src="/images/pages/tit_passMod.jpg" alt="비밀번호 변경" /></h2>
			<p class="note">
				<img src="/images/pages/txt_passMod.jpg" alt="현재 비밀번호를 입력한 후 새로 사용할 비밀번호를 입력하세요." />
			</p>
			<div class="board_write">
				<form id="frmModify" name="frmModify" method="post">
				<input type="hidden" id="Flag" name="Flag" value="PWD_C">
				<table>
					<caption>회원가입양식</caption>
					<colgroup>
						<col class="w135" />
						<col />
					</colgroup>
					<tbody>
						<tr>
							<th><label for="">현재 비밀번호</label></th>
							<td><input type="password" name="Passwd" id="Passwd" class="w130" maxlength="30" /></td>
						</tr>
						<tr>
							<th><label for="">새 비밀번호 입력</label></th>
							<td><input type="password" name="New_Passwd" id="New_Passwd" class="w130" maxlength="30" /></td>
						</tr>
						<tr>
							<th><label for="">새 비밀번호 확인</label></th>
							<td><input type="password" name="Re_Passwd" id="Re_Passwd" class="w130" maxlength="30" /></td>
						</tr>
					</tbody>
				</table>
				</form>
			</div>
			<p class="popBtn">
				<a href="javascript:chkModForm();"><img src="/images/pages/btn_pop_confirm.png" alt="확인" /></a>
				<a href="#" id="btnClose"><img src="/images/pages/btn_pop_cancel.png" alt="취소" /></a>
			</p>
		</div>
	</div></div>
