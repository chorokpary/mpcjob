// 개발 관련 기능	모음 스크립트입니다.	그 외의 경우는 추가하지 마세요. 
///// Button Control Script	: Start
$(function() {
	if (window.opener){
		$(window).load(function(){ $("body").focus(); });
	}
});
///// Button Control : End

///// Input	validate : START
// checkAll
var	global_chkall =	false;
function chekAllBtn(argTarget, argTxt){
	if (global_chkall) {
		global_chkall =	false; 
		$(argTxt).html("전체선택");
	} else {
		global_chkall =	true; 
		$(argTxt).html("전체해제");
	}

	checkAll(argTarget,global_chkall);
}

function checkAll(argTarget, argChk){
	var	$checkboxes	= $(argTarget).find(':checkbox');
	
	$checkboxes.attr('checked',	argChk);
}

function cntChecked(argTarget){
	var	chkCnt = $(argTarget).find(':checkbox:checked').length;
	
	return chkCnt;
}

function getCheckedVal(argTarget){
	var	chkVal;
	var	items =	new	Array;
	
	$("input[name='" + argTarget + "']:checkbox:checked").each(function(){
		items.push($(this).val());
	});	
	
	chkVal = items.join(',');
	
	return chkVal;
}

function getCheckboxAllVal(argTarget){
	var	chkVal;
	var	items =	new	Array;
	
	$("input[name='" + argTarget + "']:checkbox").each(function(){
		items.push($(this).val());
	});	
	
	chkVal = items.join(',');
	
	return chkVal;
}
///// Input	validate : END


///// BBS Script : START
function initPage(){
	$("#Page").val(1);
}

function goPage(p) {
	$("#frmPage").find("#Page").val(p);
	$("#frmPage").attr({action:window.location.href, method:"post",	target:"_self"}).submit();
}

function goPageAjax(p, url,	target)	{

	$("#frmPageAjax").find("#Page").val(p);

	var	param =	$("#frmPageAjax").serialize();
	// html	작업되지 않은	상태
	$.ajax({
		type: "POST"
		, url: url
		, data:	param
		, success: function(content) { 
			$(target).html(content);
		}
		, error: function(xhr, status, error) {	alert("요청하신	페이지를 찾을	수 없습니다."); /*alert(error);*/}
	});

	//$("#frmPageAjax").attr({action:window.location.href, method:"post", target:"_self"}).submit();
}

function goDown(seq, flag, fld,	fname) {
	$("#dnSeq").val(seq);
	$("#dnFlag").val(flag);
	$("#dnFld").val(fld);
	$("#dnFName").val(fname);
	$("#frmDown").attr({action:"/common/Dev/FileDown.asp", method:"post"}).submit();
}

function goDownStatic(fname) {
	window.location.href = "/common/FileDown.asp?dnFlag=STATIC&dnFName=" + encodeURIComponent(fname);
}

function goListCommon(url){
	$("#frmPage").attr({action:url,	method:"post"}).submit();
}

function goViewCommon(seq, url){
	$("#frmPage").find("#Seq").val(seq);
	$("#frmPage").attr({action:url,	method:"get"}).submit();
}


$(function() {
	$("#sKey").change(initPage);
	$("#sWord").keydown(initPage);
});
///// BBS Script : End





///// Pattern Check	: START
function IsEmail (emailStr)	{ // E-Mail	체크
	var	emailPat=/^(.+)@(.+)$/ 
	var	specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]" 
	var	validChars="\[^\\s"	+ specialChars + "\]" 
	var	firstChars=validChars 
	var	quotedUser="(\"[^\"]*\")" 
	var	ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/ 
	var	atom="(" + firstChars +	validChars + "*" + ")" 
	var	word="(" + atom	+ "|" +	quotedUser + ")" 
	var	userPat=new	RegExp("^" + word +	"(\\." + word +	")*$") 
	var	domainPat=new RegExp("^" + atom	+ "(\\." + atom	+")*$")	
	
	
	var	matchArray=emailStr.match(emailPat)	
	if (matchArray==null) {	
		//alert("@,.") 
		return false 
	} 
	var	user=matchArray[1] 
	var	domain=matchArray[2] 
	
	if (user.match(userPat)==null) { 
		//alert("id") 
		return false 
	} 
	
	var	IPArray=domain.match(ipDomainPat) 
	if (IPArray!=null) { 
		for	(var i=1;i<=4;i++) { 
			if (IPArray[i]>255)	{ 
				//alert("IP") 
				return false 
			} 
		} 
		return true	
	} 
	
	var	domainArray=domain.match(domainPat)	
	
	if (domainArray==null) { 
		//alert("domain") 
		return false 
	} 
	var	atomPat=new	RegExp(atom,"g") 
	var	domArr=domain.match(atomPat) 
	var	len=domArr.length 
	
	if (domArr[domArr.length-1].length<2 ||	
		domArr[domArr.length-1].length>3) {	
		//alert("country") 
		return false 
	} 
	
	if (domArr[domArr.length-1].length==2 && len<3)	{ 
		var	errStr="This address ends in two characters, which is a	country" 
		errStr+=" code.	Country	codes must be preceded by "	
		errStr+="a hostname	and	category (like com,	co,	pub, pu, etc.)"	
		//alert(errStr)	
		return false 
	} 
	
	if (domArr[domArr.length-1].length==3 && len<2)	{ 
		//var errStr="hostname"	
		//alert(errStr)	
		return false 
	} 
	return true; 
}

function IsSSN(jnum_01,	jnum_02) { // 주민등록번호 체크
	check_jumin	= false;
	// 잘못된 생년월일	검사.
	b_Year = (jnum_02.charAt(0)	<= "2")	? "19" : "20";
	/* 2000년도부터	성구별	번호 변경
	구별수가 2보다 작다면 1900년도	생이되고 2보다 크다면 2000년도	이후.	
	단 1800년도 생은	계산에서 제외합니다.*/
	
	b_Year += jnum_01.substr(0,	2);
	// 생년계산.
	b_Month	= jnum_01.substr(2,	2) - 1;
	// 월계산.
	b_Date = jnum_01.substr(4, 2);
	
	b_sum =	new	Date(b_Year, b_Month, b_Date);
	if ( b_sum.getYear() % 100 != jnum_01.substr(0,	2) || b_sum.getMonth() != b_Month || b_sum.getDate() !=	b_Date)	{
		return check_jumin;
		// 생년월일의 타당성을 검사하여 거짓이 있을시 에러메세지를 나타냄
	}
	
	total =	0;
	temp = new Array(13);
	for(i=1; i<=6; i++)	temp[i]	= jnum_01.charAt(i-1);
	for(i=7; i<=13;	i++) temp[i] = jnum_02.charAt(i-7);
	for(i=1; i<=12;	i++) {
		k =	i +	1;
		if(k >=	10)	k =	k %	10 + 2;
		/* 각 수와	곱할 수를 뽑아냅니다. 곱수가 만일	10보다 크거나 같다면
		계산식에 의해	2로 다시 시작하게 됩니다.	*/
		total =	total +	(temp[i] * k);
		// 각 자리수와 계산수를 곱한값을	변수 total에 누적합산시킵니다.
	}
	
	last_num = (11-	(total % 11)) %	10;
	// 마지막 계산식을	변수 last_num에 대입합니다.
	if(last_num	== temp[13]) check_jumin = true;
	// laster_num이 주민번호의마지막수와 같은면 참을 틀리면 거짓을 반환합니다.
	else check_jumin = false;
	
	return check_jumin;
}

// 모바일 형식 확인
function IsMobile(strMobile,strIdx){
	if (!strIdx){
		// 자릿수 구분 문자 없을경우
		if	(strMobile.substring(0,1) != "0" ||	strMobile.length < 10 || strMobile.length >	11 || !IsNumerics(strMobile)){
			return false;
		}
	}else{
		// 자릿수 구분 문자 있을경우
		var arrayMobile = strMobile.split(strIdx);
		
		if (arrayMobile.length != 3)		return false;
		
		if	(arrayMobile[0].substring(0,1) != "0" ||	arrayMobile[0].length != 3 || (arrayMobile[1].length < 3 && arrayMobile[1].length > 4) || arrayMobile[2].length > 4){
			return false;
		}

		if	(!IsNumerics(arrayMobile[0]) ||	!IsNumerics(arrayMobile[1]) || !IsNumerics(arrayMobile[2])){
			return false;
		}
		
		
	}
	return true;
}


function IsDate(dateStr) {
	
	//var datePat =	/^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
	var	datePat	= /^(\d{1,4})(\-)(\d{1,2})(\-)(\d{2})$/;
	var	matchArray = dateStr.match(datePat); //	is the format ok?
	
	if (matchArray == null)	{
		//alert("Please	enter date as either mm/dd/yyyy	or mm-dd-yyyy.");
		return false;
	}
	
	year = matchArray[1]; // p@rse date	into variables
	month =	matchArray[3];
	day	= matchArray[5];

	if (month <	1 || month > 12) { // check	month range
		//alert("Month must	be between 1 and 12.");
		return false;
	}
	
	if (day	< 1	|| day > 31) {
		//alert("Day must be between 1 and 31.");
		return false;
	}
	
	if ((month==4 || month==6 || month==9 || month==11)	&& day==31)	{
		//alert("Month "+month+" doesn`t have 31 days!")
		return false;
	}
	
	if (month == 2)	{ // check for february	29th
		var	isleap = (year % 4 == 0	&& (year % 100 != 0	|| year	% 400 == 0));
		if (day	> 29 ||	(day==29 &&	!isleap)) {
			//alert("February "	+ year + " doesn`t have	" +	day	+ "	days!");
			return false;
		}
	}
	return true; //	date is	valid
}

function IsDate2(dateStr) {
	if (dateStr.length != 6)	return false;
	// 잘못된 생년월일	검사.
	b_Year = "19";
	
	/* 2000년도부터	성구별	번호 변경
	구별수가 2보다 작다면 1900년도	생이되고 2보다 크다면 2000년도	이후.	
	단 1800년도 생은	계산에서 제외합니다.*/
	
	b_Year += dateStr.substr(0,	2);
	// 생년계산.
	b_Month	= dateStr.substr(2,	2) - 1;
	// 월계산.
	b_Date = dateStr.substr(4, 2);
	
	b_sum =	new	Date(b_Year, b_Month, b_Date);
	
	if ( b_sum.getYear() % 100 != dateStr.substr(0,	2) || b_sum.getMonth() != b_Month || b_sum.getDate() !=	b_Date)	{
		// 생년월일의 타당성을 검사하여 거짓이 있을시 에러메세지를 나타냄
		return false;
	}
	return true;
}

function IsNumerics(checkStr) {	// string
	var	checkOK	= "0123456789" ;
	var	isnumeric =	false ;

	for	(i = 0;	 i < checkStr.length;  i++)	{
		ch = checkStr.charAt(i)	;

		isnumeric =	false ;
		for	(j = 0;	 j < checkOK.length;  j++) {
			if (ch == checkOK.charAt(j)) {
				isnumeric =	true ;
			}
		}

		if (isnumeric == false)
			return false ;
	}
	return true;
}

/* input use numeric */
function OnlyNumber(e, decimal)	{ // onkeyup
	var	key;
	var	keychar;
	
	if (window.event) {
		key	= window.event.keyCode;
	} else if (e) {
		key	= e.which;
	} else {
		return true;
	}
	
	keychar	= String.fromCharCode(key);

	if ((key ==	null) || (key == 0)	|| (key	== 8) || (key == 9)	|| (key	== 13) || (key == 27)) {
		return true;
	} else if ((("0123456789").indexOf(keychar)	> -1)) {
		return true;
	} else if (decimal && (keychar == "."))	{
		return true;
	} else
		return false;
}


function FillNumerics(checkStr)	{ // string	return string
	var	checkOK	= "0123456789" ;
	var	rtn	= "" ;

	for	(i = 0;	 i < checkStr.length;  i++)	{
		ch = checkStr.charAt(i)	;
		for	(j = 0;	 j < checkOK.length;  j++) {
			if (ch == checkOK.charAt(j)) {
				rtn	+= ch ;
			}
		}
	}
	return rtn;
}

/* input use numeric */
function OnlyNumber_H(e, decimal)	{ // onkeyup
	var	key;
	var	keychar;
	
	if (window.event) {
		key	= window.event.keyCode;
	} else if (e) {
		key	= e.which;
	} else {
		return true;
	}
	
	keychar	= String.fromCharCode(key);

	if ((key ==	null) || (key == 0)	|| (key	== 8) || (key == 9)	|| (key	== 13) || (key == 27)) {
		return true;
	} else if ((("0123456789-").indexOf(keychar)	> -1)) {
		return true;
	} else if (decimal && (keychar == "."))	{
		return true;
	} else
		return false;
}


function FillNumerics_H(checkStr)	{ // string	return string
	var	checkOK	= "0123456789-" ;
	var	rtn	= "" ;

	for	(i = 0;	 i < checkStr.length;  i++)	{
		ch = checkStr.charAt(i)	;
		for	(j = 0;	 j < checkOK.length;  j++) {
			if (ch == checkOK.charAt(j)) {
				rtn	+= ch ;
			}
		}
	}
	return rtn;
}
///// Pattern Check	: END


/* 브라우져	체크 */
function chkBrowser(){
	var	browser;
	if(navigator.userAgent.indexOf("MSIE") != -1){
		browser	= "IE";
	}else if(navigator.userAgent.indexOf("Chrome") != -1){
		browser	= "CHROME";
	}else if(navigator.userAgent.indexOf('iPhone') != -1 ||	navigator.userAgent.indexOf('iPad')	!= -1){
		browser	= "IPHONE";
	}else if(navigator.userAgent.indexOf('Android')	!= -1 || navigator.userAgent.indexOf('Android')	!= -1){
		browser	= "ANDROID";
	}else{
		browser	= "OTHER";
	}
	
	return browser;
}

function getAjaxPage(url ,param, objTarget)	{
	$.ajax({
		type: "GET",
		url: url,
		data : param,
		dataType: "html",
		async: false,
		success: function(msg){	objTarget.html(msg);},
		error :	function(xhr, status, error){ objTarget.html("[Error] 관련 글을	찾을 수 없습니다.");}
	});
}

var global_alert_flag = false;
function limitCharactersByte(argObj,argMaxLen, argMsgTarget, argMsgType) {
	var	msgVal = argObj.val();
	var	bytesLen = 0;
	var	curMsgLen =	'';
	var	curBytesLen	= 0;
	var	realVal	= '';
	var	realLen	= 0;
	
	for(var	i =	0; i < msgVal.length; i++) {
		var	oneChar	= msgVal.charAt(i);

		if ( oneChar ==	'\n' )
			bytesLen +=	2;
		else if	( escape(oneChar).length > 4 ||	oneChar	== "°" )
			bytesLen +=	2;
		else if	( oneChar != '\r' || oneChar !=	'\n' )
			bytesLen++;
			
		if ( bytesLen <= argMaxLen ) {
			curMsgLen =	i +	1;
			curBytesLen	= bytesLen;
		}
			
	}

	if ( bytesLen >	argMaxLen && argMsgType =="L")	{
		// alert삭제하면 문자열 자르기가 정상 작동하지 않음.
		alert(argMaxLen	+ "bytes 이상의 메세지는 전송하실 수 없습니다.");
		argObj.val(msgVal.substr(0,	curMsgLen));		// 초과 입력시 초과된 만큼 잘라줌
		realLen	= curBytesLen;
	} else if ( bytesLen >	argMaxLen && argMsgType =="S")	{
		if (!global_alert_flag) { // 페이지 로딩당 최초 1회만 알림
			alert(argMaxLen	+ "bytes 이상의 메세지는 나뉘어 여러건의 SMS로 발송됩니다.");
			global_alert_flag = true;
		}
		realLen	= bytesLen;
	} else {
		realLen	= bytesLen;
	}

	argMsgTarget.html(realLen +	"/" +	argMaxLen +	" Bytes");
}

/* input에 글자수 제한*/
function limitCharacters(textid, limit,	limitid) {
	// 잆력 값	저장
	var	text = $('#'+textid).val();
	var	rtn	= false;
	
	// 입력값 길이 저장
	var	textlength = text.length;
	if(textlength >	limit) {
		// 제한 글자 길이만큼 값	재 저장
		//$('#'+textid).val(text.substr(0,limit));
		rtn	=  false;
	} else {
		rtn	=  true;
	}
	
	if (limitid)	$('#' +	limitid).html( textlength +	' /	 '+limit+ '자');
	
	return rtn;
}

/* Cookie Setting */
function setCookie(	name, value, expiredays	)
{
	var	todayDate =	new	Date();
	todayDate.setDate( todayDate.getDate() + expiredays);
	document.cookie	= name + "=" + escape( value ) + ";	path=/;	expires=" +	todayDate.toGMTString()	+ ";"
}

function getCookie(	name ) {
	var	nameOfCookie = name	+ "=";
	var	x =	0;
	while (	x <= document.cookie.length	)
	{
			var	y =	(x+nameOfCookie.length);
			if ( document.cookie.substring(	x, y ) == nameOfCookie ) {
					if ( (endOfCookie=document.cookie.indexOf( ";",	y )) ==	-1 )
							endOfCookie	= document.cookie.length;
					return unescape( document.cookie.substring(	y, endOfCookie ) );
			}
			x =	document.cookie.indexOf( " ", x	) +	1;
			if ( x == 0	)
					break;
	}
	return "";
}

function replaceAll(str,searchStr,replaceStr){
	while (str.indexOf(searchStr) != -1) {
		str	= str.replace(searchStr, replaceStr);
	}
	return str;
}


function getFlash(Path,Width,Height) {

	var	obj	= "";
	obj	= "<obj	classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='" +	Width +	"' height='" + Height +	"'>" ;
	obj	+= '<param name="movie"	value="' + Path	+ '">';
	obj	+= "<param name='allowScriptAccess'	value='always' />";
	obj	+= '<param name="quality" value="high">';
	obj	+= '<param name="wmode"	value="transparent">';
	obj	+= '<param name="allowFullScreen" value="true">';
	obj	+= '<embed src="' +	Path + '" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"	width="' + Width + '" height="'	+ Height + '" wmode="transparent" allowScriptAccess="always"></embed>';
	obj	+= '</obj>';
	
	document.write(obj);
}


// 로컬이미지 미리보기
function fileUploadPreview(thisObj,	preViewer) {
	if(!/(\.gif|\.jpg|\.jpeg|\.png)$/i.test(thisObj.value))	{
		//alert("이미지 형식의 파일을 선택하십시오");
		return;
	}

	preViewer =	(typeof(preViewer) == "object")	? preViewer	: document.getElementById(preViewer);
	var	ua = window.navigator.userAgent;

	if (ua.indexOf("MSIE") > -1) {
		var	img_path = "";
		if (thisObj.value.indexOf("\\fakepath\\") <	0) {
			img_path = thisObj.value;
		} else {
			thisObj.select();
			var	selectionRange = document.selection.createRange();
			img_path = selectionRange.text.toString();
			thisObj.blur();
		}
		preViewer.style.filter = "progid:DXImageTransform.Microsoft.AlphaImageLoader(src='fi" +	"le://"	+ img_path + "', sizingMethod='scale')";
	} else {
		preViewer.innerHTML	= "";
		var	W =	preViewer.offsetWidth;
		var	H =	preViewer.offsetHeight;
		var	tmpImage = document.createElement("img");
		preViewer.appendChild(tmpImage);

		tmpImage.onerror = function	() {
			return preViewer.innerHTML = "";
		}

		tmpImage.onload	= function () {
			if (this.width > W)	{
				this.height	= this.height /	(this.width	/ W);
				this.width = W;
			}
			if (this.height	> H) {
				this.width = this.width	/ (this.height / H);
				this.height	= H;
			}
		}
		if (ua.indexOf("Firefox/3")	> -1) {
			var	picData	= thisObj.files.item(0).getAsDataURL();
			tmpImage.src = picData;
		} else {
			tmpImage.src = "file://" + thisObj.value;
		}
	}
}


