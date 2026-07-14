$(function() {
	$('.board table tr').find('td:eq(1)').addClass('subject');
	$('.applyMng table tr').find('td:eq(0)').addClass('txt_fix');
	$('.applyMng table tr').find('td:eq(1)').addClass('txt_fixSubject');
});

$(function() {
	$(".faq dt").click(function(){
		$(this).parent().find("dd div").toggleClass("on");
		$(this).toggleClass("on");
	});
});

$(function() {
	$(".cate a img").hover(function(){
		$(this).attr("src", $(this).attr("src").replace("_off.png", "_on.png"));
	}, function(){
		if (!$(this).hasClass("on")) {
			$(this).attr("src", $(this).attr("src").replace("_on.png", "_off.png"));
		}
	});
});

// 회원가입창 오픈
function doJoin(pgFlag, seq){
	openLayer("/membership/easySignUp.asp","PageFlag="+pgFlag + "&RecrSeq=" + seq, "w668");
}

// 전체메뉴 오픈
function goFullMenu(){
	openLayer("/common/include/fullmenu.asp","", "menuAll_area");
}


// 지역검색 : 동적 Select
function getSearch_AddrGugun(){
	var url = "/common/dev/ajax/make_select_gugun.asp";
	var param = "SIDO=" + $(this).val() + "&TargetID=Search_AddrGugun&Css=selRelSite";

	$.ajax({
		type: "POST"
		, url: url
		, data: param
		, success: function(content){ 
			$("#divSearchGugun").html(content);
			$("#Search_AddrGugun").selectbox();
		}
		, error: function(xhr, status, error) {alert("[주소검색] " + error);}
	});
}

// 채용정보 검색	
function setSearch_Part(v){
	$("#Search_Part").val(v);
	$("#frmSearch_Part").submit();
}

$(function() {
    $("#Search_AddrSido").change(getSearch_AddrGugun);
})

