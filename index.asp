<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : index.asp / 메인페이지
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	if Request.ServerVariables("HTTPS")="on" Then
	Response.Redirect("http://"&Request.ServerVariables("HTTP_HOST"))
	End If

	Dim listRecord, listPage
	Dim kData, kFields, i, vNum
	Dim arrNotice, arrNoticeNum, arrPopup, arrPopupNum
	DIm arrRecr, arrRecrNum
	Dim RollCnt : RollCnt = 1

	Dim arrData, arrDataNum
	
	Call Main()
	
	Sub Main()
		
		Call getPopupList()
		Call getNoticeList()
		Call getBannerData()
		
	End Sub

	Sub getPopupList()

		
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspConfPopup_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "MAIN"
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrPopupNum 	= -1
		Else
			' Seq(0), Title(1), Width(2), Height(3)
			kFields 	= Array("Seq","Title","Width","Height")
			arrPopup 	= Rs.GetRows(,,kFields)
			arrPopupNum 	= UBound(arrPopup, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub
	
	Sub getNoticeList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspBbs_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= 1
			.Parameters("@per_page").Value 	= 5
			.Parameters("@BbsKey").Value 	= "notice"

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrNoticeNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), BbsSeq(1), Subject(2), RegDate(3), RegName(4), Hit(5), Category(6)
			kFields 	= Array("TotalCnt","BbsSeq","Subject","RegDate","RegName","Hit","Category")
			arrNotice 	= Rs.GetRows(,,kFields)
			arrNoticeNum 	= UBound(arrNotice, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub	
	
	Sub getRecrList(argPage)

		Dim Rs, j
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= argPage
			.Parameters("@per_page").Value 	= 8

			.Parameters("@sViewFlag").Value = "MAIN"
			
			Set Rs = .Execute
		End With
		
		If Not Rs.Eof Then
			RollCnt = argPage
			' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","PrjName","LogoPath","AppEDate")
			arrRecr 	= Rs.GetRows(,,kFields)
			arrRecrNum 	= UBound(arrRecr, 2)
			
			%>
			<div id="slide0<%=argPage%>" class="slide">
				<ul class="pjtList">
					<% For j = 0 To arrRecrNum %>
					<li>
						<a href="javascript:goRecrView(<%=arrRecr(1,j)%>);">
							<div class="imgs">
								<img alt="<%=arrRecr(3,j)%>" src="<%=Fn_SetDefault(arrRecr(4,j),"","/images/common/no_company.jpg",arrRecr(4,j))%>">
								<div class="bg"></div>
							</div>
						</a>
						<div class="info">
							<h4><a href="javascript:goRecrView(<%=arrRecr(1,j)%>);"><%=arrRecr(3,j)%></a></h4>
							<p class="pjt"><%=arrRecr(2,j)%></p>
							<p class="due">서류마감 : <%=FN_SetDateTimeFormat(arrRecr(5,j),"MM/DD")%></p>
						</div>
					</li>
					<% Next %>
				</ul>
			</div>
			<%
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

	Sub getBannerData()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspConfFlash_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "MAIN"
			
			' ## 회원정보
			Set Rs = .Execute

			If Rs.Eof or Rs.Bof Then
				arrDataNum 	= -1
				kData 		= False
			Else
				' Seq(0), ImgPath(1), LinkUrl(2), Sort(3)
				kFields 	= Array("Seq","ImgPath","LinkUrl","Sort")
				arrData 	= Rs.GetRows(,,kFields)
				arrDataNum 	= UBound(arrData, 2)
				kData 		= True
			End If
			
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub




%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<meta name="robots" content="index,follow"/>
	<title>엠피씨플러스 채용</title>
	<meta name="description" content="콜센터 아웃소싱(일자리창출, 가족친화인증), 컨택센터 인바운드, 아웃바운드, 교육, 컨설팅, 인재파견 등" />
	<meta name="keywords" content="mpc, 엠피씨플러스, mpc채용, mpcjob, 엠피씨, 엠피씨플러스 채용" />
        <meta property="og:type" content="website" />
	<meta property="og:title" content="엠피씨플러스 채용" />
	<meta property="og:description" content="콜센터 아웃소싱(일자리창출, 가족친화인증), 컨택센터 인바운드, 아웃바운드, 교육, 컨설팅, 인재파견 등">
	<meta property="og:url" content="http://www.mpcjob.co.kr/" />
	<meta property="og:image" content="http://www.mpcjob.co.kr/images/common/logo_800_new3.png"/>
	<link href="/common/css/main.css" rel="stylesheet" type="text/css" />
	<link type="text/css" rel="stylesheet" href="/common/js/slides/css/rcarousel.css" />
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.core.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.widget.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.rcarousel.min.js"></script>
	<script type="text/javascript" src="common/js/slides/lib/jquery.easing.1.3.js"></script>
	<script type="text/javascript" src="common/js/slides/lib/jquery.swiper-2.1.min.js"></script>
	<script type="text/javascript" src="/common/js/selectbox-min.js"></script>
	<script type="text/javascript" src="/common/js/common.js"></script>
	<script type="text/javascript" src="/common/js/common_dev.js"></script>
	<script type="text/javascript" src="/common/js/popup.js"></script>
	<script type="text/javascript" src="/common/js/jquery.form.js"></script>
	<script type="text/javascript" src="/common/js/flashObj.js"></script>
	<script type="text/javascript">
		$(window).ready( function(){
			$("#Search_AddrSido").selectbox();
			$("#Search_AddrGugun").selectbox();
		})
		
		function goRecrView(seq){
			goViewCommon(seq,"/recruit/recruit_view.asp");
		}
		
		// 팝업영역
		var getCookie = document.cookie;
		var WidthSum = 0;
	
		<% For i = 0 To arrPopupNum %>
		if( getCookie.indexOf('popup<%=arrPopup(0,i)%>') < 0 ) {
			openWin("popup.asp?seq=<%=arrPopup(0,i)%>","popup<%=arrPopup(0,i)%>","0",WidthSum,"<%=arrPopup(2,i)%>","<%=arrPopup(3,i) + 20%>",0);
			WidthSum = WidthSum + <%=arrPopup(2,i)%> + 10;
		}
		<% Next %>

	</script>

	<!--[if	IE 6]>
	<script type="text/javascript" src="/common/js/DD_belatedPNG_0.0.8a.js"></script>
	<script type="text/javascript">
		DD_belatedPNG.fix('img,#header .gnb li a:hover,#header .gnb li.on a, .pjtList li .imgs .bg, .main .bottom .inner1, .main .bottom .inner3');
		
		$(window).load(function(){ openLayer('/etc/ie6Pop.asp','','w668'); })
	</script>
	<![endif]-->
</head>

<body>
	<div id="wrap" class="main">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<!-- 플래시영역 //-->
			<div class="topRight">
			
			<!-- 20190311 포엠 banner slide -->
		    <div id="mBanner" class="banner_visual">
			<!-- slide swiper lists -->
			<div class="swiper-wrapper">
				<% If arrDataNum > -1 Then %>
					<% For i = 0 To arrDataNum %>
					<div class="swiper-slide"><A HREF="<%=arrData(2,i)%>" target="<%=Fn_SetDefault(InStr(arrData(2,i),"http://"),"0","_self","_blank")%>"><img src="<%=arrData(1,i)%>" alt="메인배너_<%=i+1%>" /></A></div>
					<% Next %>
				<% Else%>
					<img src="/images/main/banner_no.jpg" alt="메인배너" />
				<% End If %>
			</div>
		    </div>
			<% If arrDataNum > -1 Then %>
		    <div class="banner_navigation">
			<span>
			    <ul>
				<li>
				    <span class="navi_cnt">1</span>
				    <span class="bar">/</span>
				    <span class="navi_all">1</span>
				</li>
				<li>
				    <a id="fPrev" class="arrow right swiper-button-prev" href="javascript:;"><img src="images/main/banner/arrow_prev.png" alt="" /></a>
				</li>
				<li>
				    <a id="fNext" class="arrow right swiper-button-next" href="javascript:;"><img src="images/main/banner/arrow_next.png" alt="" /></a>
				</li>
			    </ul>
			</span>
		    </div>
			<% End If %>

			<script>
		    $(function(){
				//메인 슬라이드 배너
				var swiperEvent = {
					init:function(){
					var curPage = 0;
					var mySwiper = new Swiper("#mBanner",{
						slidesPerView:1,
						slidesPerGroup:1,
						speed:800,
						autoplay:5000,
						loop: false,
						prevButton: "#fPrev",
						nextButton: "#fNext",
						onSlideChangeStart: function (swiper) {
						curPage = swiper.activeIndex + 1;
						$(".banner_navigation .navi_cnt").text(curPage);
						swiper.startAutoplay();
						},
						onSlideChangeEnd: function (swiper) {
						curPage = swiper.activeIndex + 1;
						$(".banner_navigation .navi_cnt").text(curPage);
						swiper.startAutoplay();
						}
					});
					$("#fPrev").on("click",function(){ mySwiper.swipePrev(); });
					$("#fNext").on("click",function(){ mySwiper.swipeNext(); });

					var allPage = $("#mBanner .swiper-wrapper .swiper-slide").length;
					$(".navi_all").text(allPage);
					}
				};

				swiperEvent.init();

		    });
		    </script>

		    <!-- //20190311 포엠 banner slide -->

			</div>
			<div class="contents">
				<div class="clearfix section1">
					<div class="fl_l">
						<h2><img src="/images/main/stit_main_01.png" alt="직종별 채용정보" /></h2>
						<p><img src="/images/main/txt_main_01.png" alt="직종별 채용정보를 한번의 클릭으로 빠르게 찾아보세요." /></p>
					</div>
					<div class="cate fl_r">
						<form name="frmSearch_Part" id="frmSearch_Part" method="post" action="/recruit/recruit_list.asp">
						<input type="hidden" name="Search_Part" id="Search_Part" value="<%=FN_Req("GLOBAL_SEARCH_PART","")%>">
							<a href="javascript:setSearch_Part('공공기관');"><img src="/images/main/link_main_01_off.png" alt="공공기관" /></a>
							<a href="javascript:setSearch_Part('통신분야');"><img src="/images/main/link_main_02_off.png" alt="통신분야" /></a>
							<a href="javascript:setSearch_Part('금융분야');"><img src="/images/main/link_main_03_off.png" alt="금융분야" /></a>
							<a href="javascript:setSearch_Part('IT분야');"><img src="/images/main/link_main_04_off.png" alt="IT분야" /></a>
							<a href="javascript:setSearch_Part('쇼핑분야');"><img src="/images/main/link_main_05_off.png" alt="쇼핑분야" /></a>
							<a href="javascript:setSearch_Part('교육분야');"><img src="/images/main/link_main_06_off.png" alt="교육분야" /></a>
							<a href="javascript:setSearch_Part('건설분야');"><img src="/images/main/link_main_07_off.png" alt="건설분야" /></a>
							<a href="javascript:setSearch_Part('기타분야');"><img src="/images/main/link_main_08_off.png" alt="기타분야" /></a>
						</form>
					</div>
				</div>
				<div class="section2">
					<h2><img src="/images/main/stit_main_02.png" alt="지역별 채용정보" /></h2>
					<p><img src="/images/main/txt_main_02.png" alt="찾으시는 지역을 검색하세요. 검색을 통해 각 권역별 맞춤 채용정보를 알려드립니다." /></p>
					<div class="selectBox">
						<form name="frmLnbAreaSearch" id="frmLnbAreaSearch" method="post" action="/recruit/recruit_list.asp">
							<div class="inner1">
								<%
								Dim Search_SidoSql : Search_SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
								response.write " ": Call SB_MakeHTMLSelect(objDBCon, Search_SidoSql, "Search_AddrSido", "지역", "", GLOBAL_SEARCH_SIDO, "", " class=""selRelSite"" ", False)
								%>
							</div>
							<div class="inner2">
								<div id="divSearchGugun">
								<%
								Dim Search_GugunSql : Search_GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '" &  GLOBAL_SEARCH_SIDO & "' GROUP BY GUGUN ORDER BY GUGUN" 
								response.write " ": Call SB_MakeHTMLSelect(objDBCon, Search_GugunSql, "Search_AddrGugun", "지역중분류", "",  GLOBAL_SEARCH_GUGUN, "", " class=""selRelSite"" ", False)
								%>
								</div>
							</div>
							<input type="image" src="/images/main/search_loc_main.png" alt="검색" />
						</form>
					</div>
				</div>
				<div class="section3">
					<h2><img src="/images/main/stit_main_03.png" alt="HC 공지사항" /></h2>
					<ul>
					<% if arrNoticeNum = -1 Then %>
						<li>등록된 공지사항이 없습니다.</li>
					<% Else %>
						<% For i = 0 To arrNoticeNum %>
						<li><a href="/inform/notice_view.asp?Seq=<%=arrNotice(1,i)%>"><%=FN_ClearTag(arrNotice(2,i))%></a></li>
						<% Next %>
					<% End If %>
					</ul>
				</div>
				<div class="section4">
					<h2><img src="/images/main/stit_main_04.png" alt="이달의 추천공고" /></h2>
					<p class="paging" id="pages"><!-- paging //--></p>
					<div id="carousel">
						<% 
						For i = 1 To 3 
							Call getRecrList(i)
						Next
						%>
					</div>
					<form name="frmPage" id="frmPage" method="post">
					<input type="hidden" id="Page" name="Page">
					<input type="hidden" id="Seq" name="Seq">
					</form>
				</div>

				 <div class="new_bottom">
					<ul>
						<li>
						<div>
							<p class="tit">대표번호</p>
							<img src="/images/footer/icon00.png" alt="상담사 아이콘" />
							<p class="pNumber">02-3432-6100</p>
						</div>
						</li>
						<li>
						<div>
							<p class="tit">대표 이메일</p>
							<img src="/images/footer/icon01.png" alt="메일 아이콘" />
							<p class="pNumber">hr@mpc.co.kr</p>
						</div>
						</li>
						<li class="kakao">
						<div>
							<p class="tit">카카오톡 채팅상담</p>
							<img src="/images/footer/icon02.png" alt="카카오톡 아이콘" class="kakao" />
							<p class="pNumber"><a href="https://pf.kakao.com/_lItnxj/chat" class="kakao" target="_blank">MPC plus</a></p>
						</div>
						</li>
						<li>
						<div>
							<p class="tit">카카오톡 QR코드</p>
							<img src="/images/footer/qrcode.png" alt="카카오톡 QR코드" class="qrcode" />
						</div>
						</li>
						<li>
						<div>
							<p class="tit">이력서 양식 다운로드</p>
							<a href="javascript:goDown('0','CONF','ConfFormDoc','');">
							<img src="/images/footer/img_main_word.png" alt="워드 다운로드" class="docDown" />
							</a>
							<a href="javascript:goDown('0','CONF','ConfFormXls','');">
							<img src="/images/footer/img_main_excel.png" alt="엑셀 다운로드" class="docDown" />
							</a>
							<a href="javascript:goDown('0','CONF','ConfFormHwp','');">
							<img src="/images/footer/img_main_hwp.png" alt="한글 다운로드" class="docDown" />
							</a>
						</div>
						</li>
					</ul>
				</div>

			</div>
		</div>

		<!--#include virtual="/common/include/footer.asp"-->

	</div>
	<!--#include virtual="/common/include/popup.asp"-->
	<script type="text/javascript">
		<% Dim Rnd : Rnd = Second(Now()) Mod RollCnt%>
		// slides
		function generatePages() {
			var _total, i, _link;
			
			_total = $("#carousel").rcarousel("getTotalPages");
				
			for ( i = 0; i < _total; i++ ) {
				_link = $( "<a href='#'></a>" );
				
				$(_link)
					.bind("click", {page: i},
						function( event ) {
							$( "#carousel" ).rcarousel( "goToPage", event.data.page );
							event.preventDefault();
						}
					)
					.addClass( "bullet" )
					.appendTo( "#pages" );
			}
				
			// mark first page as active
			$( "a:eq(<%=Rnd%>)", "#pages" )
				.addClass( "on" )
		}

		function pageLoaded( event, data ) {
			$( "a.on", "#pages" )
				.removeClass( "on" );
			$( "a", "#pages" )
				.eq( data.page )
				.addClass( "on" );
		}
		
		$(window).load(function() {
			$("#carousel").rcarousel({
				visible: 1,
				step: 1,
				speed: 2000,
				orientation: "horizontal", // vertical / horizontal
				auto: {
					enabled: true, direction: "next", interval:6000
				},
				width: 964,
				height: 206,
				start: generatePages,
				pageLoaded: pageLoaded,
				startAtPage: <%=Rnd%>
			});
		});
	</script>
</body>
</html>