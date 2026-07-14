// JScript 파일

	//-- 스크롤링
	var scroll_pixel,div_pixel,gtpos,gbpos,loop,moving_spd;
	var top_margin = 140;
	var top_margin2 =130;
	var speed = 15;
	var speed2 = 10;
	
	function CheckScrollMove()
	{
		scroll_pixel = document.body.scrollTop;
		gtpos = document.body.scrollTop+top_margin;
		gbpos = document.body.scrollTop+top_margin2;

		if(div_id.style.pixelTop < gtpos) {
			moving_spd = (gbpos-div_id.style.pixelTop)/speed;
			div_id.style.pixelTop += moving_spd;
		}
		if(div_id.style.pixelTop > gtpos) {
			moving_spd = (div_id.style.pixelTop-gtpos)/speed;
			div_id.style.pixelTop -= moving_spd;
		}
		loop = setTimeout("CheckScrollMove()",speed2);
	}

	//-- 인원확인
	function ArrangeList(filterType) {
		var cnt = eval(fast_filter(document.getElementById("receiverList"), filterType));
		
		if (!cnt)	cnt = 0;
		$("#divTotalCnt").html(cnt);
	}


	//-- 메세지 발송
	function RegistMsg() {
		var frm = document.frmMSG;
		var cnt = 0;

		if ( Trim(frm.SendMsg.value) == "" ) { alert("메세지를 입력해주세요."); cnt++; }
		if ( Trim(frm.ReceiverList.value) == "" ) { alert("수신번호를 입력해주세요."); cnt++ }	    
		if ( cnt == 0 && confirm("메세지를 발송하시겠습니까?") ) frm.submit(); 
	}

	//-- 앞/뒤 공백 문자 제거	
	function Trim(s) { 
		s = s.replace(/^\s*/,'').replace(/\s*$/, '');
		return s;
	}


	//-- 특수문자 삽입
	function SetCha(obj) {
		var str = obj.html();
		str = replaceAll(str,"&lt;","<");
		str = replaceAll(str,"&gt;",">");
		str = replaceAll(str,"&amp;","&");
		$("#SendMsg").val($("#SendMsg").val() + str);
		doCheckByte();
	}

	// 번호삽입
	function SetNum(val) {
		$("#receiverList").val($("#receiverList").val() + val + "\r\n");
		$("#divResultLast").hide();
		
		$("#chkExist").val("N");
		$("#chkError").val("N");
	}    

	// 수신번호 마감처리
	function SetLastCharacter() {
		var str = $("#receiverList").val();
		if (str.substring(str.length-1,str.length) != "\n" && str != "" )
			$("#receiverList").val( $("#receiverList").val() + "\r\n" );
	}

	// 전화번호 필터링
	function fast_filter(obj, fillType) {

		if (obj.value == "") {
			return false;
		}

		var reVal = obj.value;
		var rePhone = '';
		var reName = '';
		var message = '';
		var countList = 0;
		var HTML = '';

		var arrayList = reVal.split("\n");
		var lengthList = arrayList.length;
		var spacer = "                ";
		var pattern = /([0-9-().― ]{1,13})[ \t]*([\W\w]*)/;
//	    var pattern = /([0-9-(). ]{9,13})[ 	]*([\W\w]*)/;

		for (var i = 0; i < lengthList; i++)
		{
			var strLine = '';
			
			// 엑셀에서 온 경우
			if ( arrayList[i].length > 13 )
				if ( arrayList[i].substring(13,14) == " " && arrayList[i].substring(9,10) != "," && arrayList[i].substring(10,11) != "," && arrayList[i].substring(11,12) != "," && arrayList[i].substring(12,13) != ","  )
					arrayList[i] = arrayList[i].substring(0,15) + "	" + arrayList[i].substring(14,arrayList[i].length);
			
			arrayList[i] = arrayList[i].replace(/[	]/g,",");


			row = pattern.exec(Trim(arrayList[i]));
			if (!row) continue;

//		    rePhone = row[1].replace(/[-\(\) ]/g, "");
			rePhone = row[1].replace(/[-\(\)\.― ]/g, "");
			
			if ( fillType == 'A' ) {
				if ( (rePhone.substring(0,1) != "0") && (rePhone.substring(0,2) == "10" || rePhone.substring(0,2) == "16" || rePhone.substring(0,2) == "17" || rePhone.substring(0,2) == "18" || rePhone.substring(0,2) == "11" || rePhone.substring(0,2) == "19") ) rePhone = "0" + rePhone;
			}
			
			//reName = row[2];
			reName = row[2].replace(/[?]/g, "");

			var arrMerge = reName.split(",");
			var mergeCount = arrMerge.length;
			var m_name = '';

			if (mergeCount >= 2 || reName != "")
				rePhone = rePhone + ",";

			if (mergeCount == 2)
				m_name = Trim(arrMerge[1]);
			else if (mergeCount < 2 )
				m_name = Trim(reName);


			var nextSpace = "";
			var rtnCount = 0;

			rtnCount = SetSpaceString("             ", 0, rePhone);
			strLine = rtnVal;
			rtnVal = "";

			rtnCount = SetSpaceString("          ", rtnCount, m_name);
			strLine = strLine + rtnVal + "\r\n";
			rtnVal = "";

			HTML += strLine;
			
			
			countList++;
		}
		obj.value = HTML.substr(0, HTML.length - 2);
		
		return countList;

	}

	function SetSpaceString(spacer, cut, str)
	{
		var curByte = GetByte(str);

		if (cut < 0 && spacer.length >= (-cut))
			spacer = spacer.substr(0, spacer.length + cut);

		if (spacer.length > curByte && str.length > 0)
			rtnVal = Trim(str) + spacer.substr(0, spacer.length - curByte);
		else
			rtnVal = str;
			
		return spacer.length - curByte;
	}

	function GetByte(str)
	{
		var bytesLen = 0;

		for(var i = 0; i < str.length; i++)
		{
			var oneChar = str.charAt(i);

			if (escape(oneChar) =='%0D') {
			} else 
			if ( escape(oneChar).length > 4 || oneChar == "°" || oneChar == "¿" || oneChar.charCodeAt(0) > 127 )
				bytesLen += 2;
			else if ( oneChar != '\r' || oneChar != '\n' )
				bytesLen++;
		}

		return bytesLen;
	}
	
	/* 4M 추가 */
	// 수신자번호 배열화
	function MakeArray_RecvList(){
	
		var recvList = $("#receiverList").val();
		recvList = recvList.replace(/\r/gi, "\n");
		recvList = recvList.replace(/\n+/gi, "\n");
		
		var ArrayList = new Array();
		
		ArrayList = recvList.split("\n");
		
		return ArrayList
	}
	
	// 수신자번호 중복 검사
	function CheckRecvExist(){
		var count = 0;
		
		var arrRecvList = new Array();
		var arrNewList = new Array();
		var hash = new Object();
		
		arrRecvList = MakeArray_RecvList();
		
		for (var i=0; i<arrRecvList.length; i++){
			if (hash[arrRecvList[i].toLowerCase().substring(0,11)] != 1) {
				arrNewList = arrNewList.concat(arrRecvList[i]);
				hash[arrRecvList[i].toLowerCase().substring(0,11)] = 1
			} else { 
				count++; 
			}
		}
		
		if (count >0) {
			$("#receiverList").val(arrNewList.join("\r\n"));
			alert(count + "개의 중복을 찾아서 정리하였습니다.");
		} else {
			alert("중복된 번호가 없습니다.");
		}
		$("#chkExist").val("Y");
	}
	
	// 수신자번호 오류 검사
	function CheckRecvError(){
		var count = 0;
		
		var arrRecvList = new Array();
		var arrNewList = new Array();
		var hash = new Object();
		var RecvNum = "";
		
		arrRecvList = MakeArray_RecvList();
		
		for (var i=0; i<arrRecvList.length; i++){
			RecvNum = $.trim(arrRecvList[i].toLowerCase());
			if (RecvNum.indexOf(",") > 0)
				RecvNum = RecvNum.substring(0,RecvNum.indexOf(","));
				
			if (IsMobile(RecvNum)) {
				arrNewList = arrNewList.concat(arrRecvList[i]);
			} else { 
				count++; 
			}
		}
		
		if (count >0) {
			$("#receiverList").val(arrNewList.join("\r\n"));
			alert(count + "개의 오류를 찾아서 정리하였습니다.");
		} else {
			alert("입력하신 번호에 오류가 없습니다");
		}		
		$("#chkError").val("Y");
	}