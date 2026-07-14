<%
'____________________________________________________________________________________
'
' * Discription : SendForm.asp / SMS 발송폼
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-31 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim TotalCnt : TotalCnt = 0
	Dim tt

	Dim Flag, Seq
	
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","")
		Flag = FN_Req("Flag","")
		
	End Sub
	
	Sub getList()
	
		If Seq <> "" Then
	
			Dim Rs
			
			With objCmd
				.ActiveConnection = objDBCon
				.CommandType = 4
				.CommandText ="uspUser_Info"
				.Parameters.Refresh
				.Parameters("@FLAG").Value 	= Flag & "_SEL"
				.Parameters("@StrSeq").Value = Replace(Seq," ","")
				
				' ## 회원정보
				Set Rs = .Execute
				
				If Not Rs.Eof Then
				     Do Until Rs.Eof
				     	Response.Write Left(Rs("Mobile") & ",              ",13)
				     	Response.Write Rs("UsrName")
				     	Response.Write Chr(13) + Chr(10)
				     	Rs.MoveNext
				     Loop
				End If
	
				Rs.Close
				Set Rs = Nothing
			End With
		End If
	End Sub
%>
<link href="/common/js/jquery-ui_css/jquery-ui-1.8.21.custom.css" rel="stylesheet" type="text/css" />
<link href="/common/Dev/SMS/css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/common/js/jquery-ui-1.8.21.custom.min.js"></script>
<script type="text/javascript" src="/common/Dev/SMS/js/ASPMessageCommon.js"></script>

<script type="text/javascript">
	function chkForm(){
		ArrangeList("M");
		
		if (!$("#receiverList").val()){
			alert("받는분 전화번호를 입력해주세요.");
			$("#receiverList").focus();
			return false;
		}

		/*
		if ($("#chkExist").val()=="N"){
			alert("[중복번호제거]를 진행해주세요.");
			return false;
		}
		
		if ($("#chkError").val()=="N"){
			alert("[번호오류검사]를 진행해주세요.");
			return false;
		}
		*/

		if (!$("#SendMsg").val()){
			alert("전송 메시지를 입력해주세요.");
			$("#SendMsg").focus();
			return false;
		}
		
        if ($("#SendMsg").val().indexOf("{이름}") > -1 && !$("input[name=UseMerge]:checked").val() ) {
            if ( confirm("메세지 내용에 `{이름}` 이 포함되어 있습니다.\n머지 기능을 사용하시겠습니까? ") )
            	$("#UseMerge").attr("checked",true);
        }

		if (!$("#callbackPhoneNo").val()){
			alert("보내는 사람 전화번호를 입력해주세요.");
			$("#callbackPhoneNo").focus();
			return false;
		}
        
		//$("#Flag").val("ANS");
		$("#frmSms").attr({action:"/common/Dev/SMS/SendProc_n.asp", method:"post"});
		
		return true;
	}
	
	function doSearch(seq){
		// 전화번호 검색
		if (!$("#SearchKey").val()) {
			alert("검색어를 입력해 주세요.");
			$("#SearchKey").focus();
			return;
		}
		if ($("#SearchKey").val().length < 2) {
			alert("검색어는 2글자 이상 입력해 주세요.");
			$("#SearchKey").focus();
			return;
		}

		var url = "/common/Dev/SMS/Ajax_SearchList.asp";
		var param = "SearchKey=" + $("#SearchKey").val();

		SetLastCharacter();
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				$("#divResultSearch").html(result);
				$("#divResultSearch").show();
			}
			, error: function(xhr, status, error) {alert(error);}
		});
	}
	
	function doShowLastRecv(){
		// 최근발신 번호
		var url = "/common/Dev/SMS/Ajax_LastRecvList.asp";
		var param = "";
		
		SetLastCharacter();
			
		$.ajax({
			type: "POST"
			, url: url
			, data: param
			, success: function(result){ 
				$("#divResultLast").html(result);
				$("#divResultLast").show();
			}
			, error: function(xhr, status, error) {alert(error);}
		});
	}
	
	function doRecvCount(){
		ArrangeList("M");
	}
	
	function doRecvCheckExist(){
		ArrangeList("A");	// 번호 정리
		CheckRecvExist();	// 중복 체크
		ArrangeList("A");	// 총인원 계산
	}
	
	function doRecvCheckErr(){
		ArrangeList("A");	// 번호 정리
		CheckRecvError();	// 번호 에러체크
		ArrangeList("A");	// 총인원 계산
	}
	
	function doRecvClear(){
		if ( confirm("수신 리스트를 삭제하시겠습니까?") ) {
			$("#receiverList").val("");
		}
	}
		
	function doCheckByte(){
		var limit = 90;
		if ($("input[name=MessageType]:checked").val() == "L"){
			limit = 2000;
		}else{
			limit = 90;
		}
		limitCharactersByte($("#SendMsg"),limit,$("#divLimitTxt"),$("input[name=MessageType]:checked").val());
	}
	
	function doShowChar(flag){
		if (flag==1){
			$("#btnCha1").addClass("on");
			$("#btnCha2").removeClass("on");
			$("#divCha1").show();
			$("#divCha2").hide();
		
		}else{
			$("#btnCha1").removeClass("on");
			$("#btnCha2").addClass("on");
			$("#divCha1").hide();
			$("#divCha2").show();
		}
	}

	function chkSearchEnter() {
		if (window.event.keyCode == 13) {
			doSearch();
			return false;
		}
	}
	
	function doCheckInit() {
		$("#chkExist").val("N");
		$("#chkError").val("N");
	}
	
	function init(){
		doCheckByte();
		doShowChar(1);
		$("#divResultSearch").hide();
		$("#divResultLast").hide();
	}

	$(document).ready(function() {

		$("#ResDate").datepicker({
			date: $(this).val()
			, current: $(this).val()
			, showOn: "focus"
		});

		// ### Button Control
		$("#btnSearch").click(doSearch);
		$("#btnLastRecv").click(doShowLastRecv);
		
		$("#btnCount").click(doRecvCount);
		$("#btnClear").click(doRecvClear);
		$("#btnChkExist").click(doRecvCheckExist);
		$("#btnChkError").click(doRecvCheckErr);

		// ### Layer Control
		$("#divResultSearch").live("mouseout", function(){ $(this).hide(); });
		$("#divResultSearch ul").live("mouseover", function(){ $("#divResultSearch").show(); });
		$("#divResultLast").live("mouseout", function(){ $(this).hide(); });
		$("#divResultLast ul").live("mouseover", function(){ $("#divResultLast").show(); });

		
		// ### Input Control
		$("#SearchKey").keydown(chkSearchEnter);
		$("#receiverList").blur(SetLastCharacter);
		$("#receiverList").keyup(doCheckInit);
		

		
		$("input[name=MessageType]").click(doCheckByte);
		$("#SendMsg").keyup(doCheckByte);
		//$("#SendMsg").blur(doCheckByte);

		$("#callbackPhoneNo").keypress(function(e){ return OnlyNumber(e,false); });
		$("#callbackPhoneNo").keyup(function(){ $(this).val( FillNumerics($(this).val())); });

	});
	
	$(window).load(init);

</script>
<div class="sms_in">
	<form id="frmSms" name="frmSms" onsubmit="javascript: return chkForm();">
	<input type="hidden" id="chkExist" name="chkExist" value="N">
	<input type="hidden" id="chkError" name="chkError" value="N">
	<div class="phoneN">
		<h3>1. 받는분 전화번호 입력</h3>
		<div class="inner">
			<div class="inputArea">
				<input type="text" name="SearchKey" id="SearchKey" class="w85" />
				<span class="btns btn_in"><a href="#" id="btnSearch">검색</a></span>
				<span class="btns btn_b"><a href="#" id="btnLastRecv">최근발신번호</a></span>
				<div id="divResultSearch"><!-- 회원검색 목록 //--></div>
				<div id="divResultLast"><!-- 최근발신번호 목록 //--></div>
			</div>
			<div class="phoneList">
				<h4 class="clearfix"><span class="phone">전화번호</span><span class="name">이름</span></h4>
				<!-- 수신자 연락처 //-->
				<textarea id="receiverList" name="receiverList" rows="8" cols="40"><% Call getList() %></textarea>
			</div>
			<p class="total">총인원 : <span id="divTotalCnt"><%=TotalCnt%></span> 명</p>
			<div class="btnArea">
				<span class="btns btn_b"><a href="#" id="btnCount">인원확인</a></span>
				<span class="btns btn_b"><a href="#" id="btnClear">전화번호삭제</a></span>
				<span class="btns btn_b"><a href="#" id="btnChkExist">중복번호제거</a></span>
				<span class="btns btn_b"><a href="#" id="btnChkError">번호오류검사</a></span>
			</div>
		</div>
	</div>
	<div class="writeSms">
		<h3>2. 전송 메시지 작성</h3>
		<div class="inner">
			<div class="inputArea">
				<input type="radio" name="MessageType" id="MessageType_S" value="S" checked /><label for="MessageType_S">SMS</label>
				<input type="radio" name="MessageType" id="MessageType_L" value="L" /><label for="MessageType_L">LMS</label>
			</div>
			<div class="clearfix">
				<div class="fl_l inputSms">
					<!-- 문자영역 //-->
					<textarea name="SendMsg" id="SendMsg" rows="8" cols="36"></textarea>
					
					<p class="byte clearfix">
						<span id="divLimitTxt" class="fl_l"><!-- Byte 체크 //--></span>
						<span class="fl_r"><span onclick="SetCha($(this));" style="cursor:pointer;">{이름}</span> <input type="checkbox" name="UseMerge" id="UseMerge" value="Y" /><label for="UseMerge">머지사용</label></span>
					</p>
				</div>
				<div class="fl_r ref">
					<div class="titles clearfix">
					  <a href="javascript:doShowChar(1);" id="btnCha1">얼굴문자</a>
					  <a href="javascript:doShowChar(2);" id="btnCha2">특수문자</a>
					</div>
					<div class="example">
						<div id="divCha1" class="inner">
							<ul>
								<li><a href="#" onclick="SetCha($(this));">>.<</a></li>
								<li><a href="#" onclick="SetCha($(this));">>^.^<</a></li>
								<li><a href="#" onclick="SetCha($(this));"><-.-></a></li>
								<li><a href="#" onclick="SetCha($(this));">>-.-<</a></li>
								<li><a href="#" onclick="SetCha($(this));">(*.*)</a></li>
								<li><a href="#" onclick="SetCha($(this));">(^.^)</a></li>
								<li><a href="#" onclick="SetCha($(this));">(__)</a></li>
								<li><a href="#" onclick="SetCha($(this));">(-.-)</a></li>
								<li><a href="#" onclick="SetCha($(this));">($.$)</a></li>
								<li><a href="#" onclick="SetCha($(this));">⊙.⊙</a></li>
								<li><a href="#" onclick="SetCha($(this));">⊙⊙;</a></li>
								<li><a href="#" onclick="SetCha($(this));">(⊙⊙)</a></li>
								<li><a href="#" onclick="SetCha($(this));">⊙-⊙</a></li>
								<li><a href="#" onclick="SetCha($(this));">⊙^^⊙</a></li>
								<li><a href="#" onclick="SetCha($(this));">^.^</a></li>
								<li><a href="#" onclick="SetCha($(this));">z(-.-)v</a></li>
								<li><a href="#" onclick="SetCha($(this));">-.-ㆀ</a></li>
								<li><a href="#" onclick="SetCha($(this));">(˘ㅇ˘)</a></li>
								<li><a href="#" onclick="SetCha($(this));">~.~</a></li>
								<li><a href="#" onclick="SetCha($(this));">*^^*</a></li>
								<li><a href="#" onclick="SetCha($(this));">~.^</a></li>
								<li><a href="#" onclick="SetCha($(this));">(@@)</a></li>
								<li><a href="#" onclick="SetCha($(this));">@@</a></li>
								<li><a href="#" onclick="SetCha($(this));">^^;</a></li>
								<li><a href="#" onclick="SetCha($(this));">(★.★)</a></li>
								<li><a href="#" onclick="SetCha($(this));">^^..</a></li>
								<li><a href="#" onclick="SetCha($(this));">d(^.^)b</a></li>
								<li><a href="#" onclick="SetCha($(this));">(*|*)</a></li>
								<li><a href="#" onclick="SetCha($(this));">≥∇≤</a></li>
								<li><a href="#" onclick="SetCha($(this));">$.$</a></li>
								<li><a href="#" onclick="SetCha($(this));">^^ㆀ</a></li>
								<li><a href="#" onclick="SetCha($(this));">(n . n)</a></li>
								<li><a href="#" onclick="SetCha($(this));">(X.X)</a></li>
								<li><a href="#" onclick="SetCha($(this));">ㅡ.ㅡ</a></li>
								<li><a href="#" onclick="SetCha($(this));">ご..ご</a></li>
								<li><a href="#" onclick="SetCha($(this));">ㅡㅡㆀ</a></li>
								<li><a href="#" onclick="SetCha($(this));">~^.^~</a></li>
								<li><a href="#" onclick="SetCha($(this));">*(ㅡㅡ)*</a></li>
								<li><a href="#" onclick="SetCha($(this));">^--^</a></li>
								<li><a href="#" onclick="SetCha($(this));">ㅠ.ㅠ</a></li>
								<li><a href="#" onclick="SetCha($(this));">ㅜ.ㅜ</a></li>
								<li><a href="#" onclick="SetCha($(this));">-.-</a></li>
								<li><a href="#" onclick="SetCha($(this));">^__^</a></li>
								<li><a href="#" onclick="SetCha($(this));">^5^</a></li>
								<li><a href="#" onclick="SetCha($(this));">^__^ㆀ</a></li>
								<li><a href="#" onclick="SetCha($(this));">^8^</a></li>
								<li><a href="#" onclick="SetCha($(this));">&amp;(**)&amp;</a></li>
								<li><a href="#" onclick="SetCha($(this));">^ㅡㅡ^</a></li>
								<li><a href="#" onclick="SetCha($(this));">[**]</a></li>
								<li><a href="#" onclick="SetCha($(this));">[^^]</a></li>
								<li><a href="#" onclick="SetCha($(this));">[__]</a></li>
								<li><a href="#" onclick="SetCha($(this));">[--]</a></li>
								<li><a href="#" onclick="SetCha($(this));">[ㅡㅡ]</a></li>
								<li><a href="#" onclick="SetCha($(this));">+(^^)+</a></li>
								<li><a href="#" onclick="SetCha($(this));">(=.=)</a></li>
								<li><a href="#" onclick="SetCha($(this));"><=.=></a></li>
								<li><a href="#" onclick="SetCha($(this));"><^5^></a></li>
								<li><a href="#" onclick="SetCha($(this));">($.$)v</a></li>
							</ul>
						</div>
						<div id="divCha2" class="inner">
							<ul>
								<li><a href="#" onclick="SetCha($(this));">☆</TD>
								<li><a href="#" onclick="SetCha($(this));">★</a></li>
								<li><a href="#" onclick="SetCha($(this));">○</a></li>
								<li><a href="#" onclick="SetCha($(this));">●</a></li>
								<li><a href="#" onclick="SetCha($(this));">◎</a></li>
								<li><a href="#" onclick="SetCha($(this));">◇</a></li>
								<li><a href="#" onclick="SetCha($(this));">◆</a></li>
								<li><a href="#" onclick="SetCha($(this));">□</a></li>
								<li><a href="#" onclick="SetCha($(this));">■</a></li>
								<li><a href="#" onclick="SetCha($(this));">△</a></li>
								<li><a href="#" onclick="SetCha($(this));">▲</a></li>
								<li><a href="#" onclick="SetCha($(this));">▽</a></li>
								<li><a href="#" onclick="SetCha($(this));">▼</a></li>
								<li><a href="#" onclick="SetCha($(this));">→</a></li>
								<li><a href="#" onclick="SetCha($(this));">←</a></li>
								<li><a href="#" onclick="SetCha($(this));">↑</a></li>
								<li><a href="#" onclick="SetCha($(this));">↓</a></li>
								<li><a href="#" onclick="SetCha($(this));">↔</a></li>
								<li><a href="#" onclick="SetCha($(this));">◁</a></li>
								<li><a href="#" onclick="SetCha($(this));">◀</a></li>
								<li><a href="#" onclick="SetCha($(this));">▷</a></li>
								<li><a href="#" onclick="SetCha($(this));">▶</a></li>
								<li><a href="#" onclick="SetCha($(this));">♤</a></li>
								<li><a href="#" onclick="SetCha($(this));">♠</a></li>
								<li><a href="#" onclick="SetCha($(this));">♡</a></li>
								<li><a href="#" onclick="SetCha($(this));">♥</a></li>
								<li><a href="#" onclick="SetCha($(this));">※</a></li>
								<li><a href="#" onclick="SetCha($(this));">♧</a></li>
								<li><a href="#" onclick="SetCha($(this));">♣</a></li>
								<li><a href="#" onclick="SetCha($(this));">⊙</a></li>
								<li><a href="#" onclick="SetCha($(this));">◈</a></li>
								<li><a href="#" onclick="SetCha($(this));">▣</a></li>
								<li><a href="#" onclick="SetCha($(this));">◐</a></li>
								<li><a href="#" onclick="SetCha($(this));">◑</a></li>
								<li><a href="#" onclick="SetCha($(this));">▒</a></li>
								<li><a href="#" onclick="SetCha($(this));">♨</a></li>
								<li><a href="#" onclick="SetCha($(this));">▤</a></li>
								<li><a href="#" onclick="SetCha($(this));">▥</a></li>
								<li><a href="#" onclick="SetCha($(this));">▨</a></li>
								<li><a href="#" onclick="SetCha($(this));">▧</a></li>
								<li><a href="#" onclick="SetCha($(this));">▦</a></li>
								<li><a href="#" onclick="SetCha($(this));">▩</a></li>
								<li><a href="#" onclick="SetCha($(this));">☏</a></li>
								<li><a href="#" onclick="SetCha($(this));">☎</a></li>
								<li><a href="#" onclick="SetCha($(this));">¶</a></li>
								<li><a href="#" onclick="SetCha($(this));">☜</a></li>
								<li><a href="#" onclick="SetCha($(this));">☞</a></li>
								<li><a href="#" onclick="SetCha($(this));">†</a></li>
								<li><a href="#" onclick="SetCha($(this));">‡</a></li>
								<li><a href="#" onclick="SetCha($(this));">↕</a></li>
								<li><a href="#" onclick="SetCha($(this));">↗</a></li>
								<li><a href="#" onclick="SetCha($(this));">↙</a></li>
								<li><a href="#" onclick="SetCha($(this));">↖</a></li>
								<li><a href="#" onclick="SetCha($(this));">↘</a></li>
								<li><a href="#" onclick="SetCha($(this));">♩</a></li>
								<li><a href="#" onclick="SetCha($(this));">♪</a></li>
								<li><a href="#" onclick="SetCha($(this));">♬</a></li>
								<li><a href="#" onclick="SetCha($(this));">㈜</a></li>
								<li><a href="#" onclick="SetCha($(this));">㏇</a></li>
								<li><a href="#" onclick="SetCha($(this));">™</a></li>
								<li><a href="#" onclick="SetCha($(this));">℡</a></li>
								<li><a href="#" onclick="SetCha($(this));">㏂</a></li>
								<li><a href="#" onclick="SetCha($(this));">㏘</a></li>
								<li><a href="#" onclick="SetCha($(this));">ⅰ</a></li>
								<li><a href="#" onclick="SetCha($(this));">ⅱ</a></li>
								<li><a href="#" onclick="SetCha($(this));">ⅲ</a></li>
								<li><a href="#" onclick="SetCha($(this));">ⅳ</a></li>
								<li><a href="#" onclick="SetCha($(this));">Ⅰ</a></li>
								<li><a href="#" onclick="SetCha($(this));">Ⅱ</a></li>
								<li><a href="#" onclick="SetCha($(this));">Ⅲ</a></li>
								<li><a href="#" onclick="SetCha($(this));">Ⅳ</a></li>
								<li><a href="#" onclick="SetCha($(this));">×</a></li>
								<li><a href="#" onclick="SetCha($(this));">∑</a></li>
								<li><a href="#" onclick="SetCha($(this));">⇔</a></li>
								<li><a href="#" onclick="SetCha($(this));">⇒</a></li>
								<li><a href="#" onclick="SetCha($(this));">♀</a></li>
								<li><a href="#" onclick="SetCha($(this));">♂</a></li>
								<li><a href="#" onclick="SetCha($(this));">∞</a></li>
								<li><a href="#" onclick="SetCha($(this));">㉮</a></li>
								<li><a href="#" onclick="SetCha($(this));">㉯</a></li>
								<li><a href="#" onclick="SetCha($(this));">㉰</a></li>
								<li><a href="#" onclick="SetCha($(this));">½</a></li>
								<li><a href="#" onclick="SetCha($(this));">⅓</a></li>
								<li><a href="#" onclick="SetCha($(this));">⅔</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎖</a></li>
								<li><a href="#" onclick="SetCha($(this));">ℓ</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎘</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎣</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎤</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎥</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎜</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎝</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎞</a></li>
								<li><a href="#" onclick="SetCha($(this));">㎡</a></li>
								<li><a href="#" onclick="SetCha($(this));">®</a></li>
								<li><a href="#" onclick="SetCha($(this));">㏂</a></li>
								<li><a href="#" onclick="SetCha($(this));">㏘</a></li>
								<li><a href="#" onclick="SetCha($(this));">™</a></li>
								<li><a href="#" onclick="SetCha($(this));">℡</a></li>
							</ul>
						</div>
					</div>
				</div>
			</div>
		</div>
		<h3>3. 예약발송</h3>
		<div class="inner send clearfix">
			<div class="fl_l inputArea">
				<input type="radio" name="ReservedChk" id="ReservedChk_N" value="N" checked /><label for="ReservedChk_N">즉시발송</label>
				<input type="radio" name="ReservedChk" id="ReservedChk_Y" value="R" /><label for="ReservedChk_Y">예약발송</label>
				<p>
					<input type="text" name="ResDate" id="ResDate" value="<%=Date()%>" maxlength="10" readonly  />
					<select name="ResHour" id="ResHour">
						<% For tt = 0 TO 23 %>
						<option value="<%=Right("0" & tt,2)%>" <%=Fn_SetDefault(tt,Hour(Now())," selected","")%>><%=Right("0" & tt,2)%></option>
						<% Next %>
					</select> 시 
					<select name="ResMin" id="ResMin">
						<% For tt = 0 TO 59 %>
						<option value="<%=Right("0" & tt,2)%>" <%=Fn_SetDefault(tt,Minute(Now())," selected","")%>><%=Right("0" & tt,2)%></option>
						<% Next %>
					</select> 분
				</p>
			</div>
			<div class="fl_r">
				<label for="">보내는 사람</label>
				<input type="text" name="callbackPhoneNo" id="callbackPhoneNo" value="<%=GLOBAL_SMS_CALLBACK%>" style="ime-mode:disabled;" maxlength="15" />
				<span class="btns btn_gr"><button type="submit">문자보내기</button></span>
			</div>
		</div>
	</div>
	</form>
</div>
