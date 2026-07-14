/* ## 팝업 컨트롤 : Window Popup ## */
function openWin(url,target,t,l,w,h,s){
	return window.open(url,target,"top="+t+",left="+l+",width="+w+",height="+h+",scrollbars="+s+",toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0");
}
	
function closeWin(){
	self.close();	
}

/* ## 팝업 컨트롤 : Layer Popup ## */
function openLayer(url,param, cls){
	// html 작업되지 않은 상태
	$.ajax({
		type: "POST"
		, url: url
		, data: param
		, success: function(content) { 
			$("#divPopArea .popBox").html(content);
			// 팝업이 열려있지 않을 경우만 열기 실행
			if ($("#divPopArea").css("display")=="none"){
				//$("#divPopArea .popBox").toggle(function(){
					$("#divPopArea .popupbg").css({"background": "#000000"});
					$("#divPopArea .popupbg").css({"opacity": "0.7"});
					$("#divPopArea .popBox").addClass(cls);
					$("#divPopArea").fadeIn();
					$("select:not('#divPopArea .popBox select')").hide();
					reposLayer();
				//});
			}
		}
		, error: function(xhr, status, error) { alert("요청하신 페이지를 찾을 수 없습니다."); /*alert(error);*/}
	});
}

function closeLayer(){
	$("#divPopArea").fadeOut();
	$("select:not('.selRelSite')").show();
	
	// 입사지원 페이지에서 창을 닫을 경우에만 reload (로그인 상태가 변동되기때문)
	if ($("#divPopArea .popBox").find(".recruitComplete").length == 1){
		window.location.reload();
	}
}

function reposLayer(){
	if ($("#divPopArea").css("display")!="none"){
		var top_m = -100;
		
		$("#divPopArea").css("height",($(window).height() + $(window).scrollTop()) + "px");  
		$("#divPopArea").css("width",($(window).width() + $(window).scrollLeft()) + "px");

		$("#divPopArea .popupbg").css("height",$("#divPopArea").height());
		
//		alert($("#footer").height());
		if ( $("#divPopArea .popBox").height() + $(window).scrollTop() + top_m < $("#wrap").height()) {
			if ($(window).height() > $('#divPopArea .popBox').height()) {
				$("#divPopArea .popBox").css("top", ( ($(window).height() - $("#divPopArea .popBox").height() ) / 2+$(window).scrollTop()) + top_m + "px");
			}else{
				$("#divPopArea .popBox").css("top",($(window).scrollTop() + top_m) + "px");
			}
		}
		if ( $('#divPopArea .popBox').width() + $(window).scrollLeft() < $("#wrap").width()) {
			if ($(window).width() > $("#divPopArea .popBox").width()) {
				$("#divPopArea .popBox").css("left", ( $(window).width() - $("#divPopArea").width() ) / 2+$(window).scrollLeft() + "px");
			}else{
				$("#divPopArea .popBox").css("left",$(window).scrollLeft() + "px");
			}
		}
	}
}


/* ## 이벤트 컨트롤 ## */

$(function() {
	// ++ window pop
	$("#btnClose").click(closeWin);
	
	// ++ layer pop
	$(window).scroll(reposLayer);
	$(window).resize(reposLayer);
	$(window).load(function(){	
		$('#divPopArea').hide();
	});

	$("#divPopArea .popupbg").click(closeLayer);
	$("#divPopArea").find("#btnClose").live("click",closeLayer);
	$(".menuAll_open").click(function(){
		openLayer("url","param","");
	});
});