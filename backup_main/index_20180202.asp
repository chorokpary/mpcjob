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
	Dim listRecord, listPage
	Dim kData, kFields, i, vNum
	Dim arrNotice, arrNoticeNum, arrPopup, arrPopupNum
	DIm arrRecr, arrRecrNum
	Dim RollCnt : RollCnt = 1
	
	Call Main()
	
	Sub Main()
		
		Call getPopupList()
		Call getNoticeList()
		
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



%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
	<meta name="robots" content="index,follow"/>
	<title>한국코퍼레이션 채용</title>
	<meta name="description" content="콜센터 아웃소싱(국내 업계최초, 코스닥상장, 일자리창출, 가족친화인증), 컨택센터 인바운드, 아웃바운드, 교육, 컨설팅, 인재파견 등" />
	<meta name="keywords" content="mpc, 한국코퍼레이션, mpc채용, mpcjob, 엠피씨, 한국코퍼레이션 채용" />
        <meta property="og:type" content="website" />
	<meta property="og:title" content="한국코퍼레이션 채용" />
	<meta property="og:description" content="콜센터 아웃소싱(국내 업계최초, 코스닥상장, 일자리창출, 가족친화인증), 컨택센터 인바운드, 아웃바운드, 교육, 컨설팅, 인재파견 등">
	<meta property="og:url" content="http://www.mpcjob.co.kr/" />
	<meta property="og:image" content="http://www.mpcjob.co.kr/images/common/logo_800_new2.png"/>
	<link href="/common/css/main.css" rel="stylesheet" type="text/css" />
	<link type="text/css" rel="stylesheet" href="/common/js/slides/css/rcarousel.css" />
	<script type="text/javascript" src="/common/js/jquery-1.8.0.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.core.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.widget.min.js"></script>
	<script type="text/javascript" src="/common/js/slides/lib/jquery.ui.rcarousel.min.js"></script>
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
			<div class="topRight"><script type="text/javascript"> getFlash("/images/flash/main.swf", 641, 341); </script></div>
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
				<div class="bottom clearfix">
					<div class="inner1">
						<h2><a href="/corp/center.asp"><img src="/images/main/stit_main_05.png" alt="한국코퍼레이션 센터를 찾으시나요? 보다 편리한 전국 한국코퍼레이션 센터를 찾아보세요." /></a></h2>
					</div>
					<div class="inner2">
						<h2><img src="/images/main/stit_main_06.png" alt="이력서 양식 다운로드" /></h2>
						<p> 
							<a href="javascript:goDown('0','CONF','ConfFormDoc','');"><img src="/images/main/img_main_word.png" alt="워드파일" /></a>
							<a href="javascript:goDown('0','CONF','ConfFormXls','');"><img src="/images/main/img_main_excel.png" alt="엑셀파일" /></a>
							<a href="javascript:goDown('0','CONF','ConfFormHwp','');"><img src="/images/main/img_main_hwp.png" alt="한글파일" /></a>
						</p>
					</div>
					<div class="inner3">
						<h2><img src="/images/main/stit_main_07.png" alt="한국코퍼레이션 채용문의" /></h2>
						<p><img src="/images/main/txt_main_07.png" alt="tel : 02-3432-6100 / phone : 010-9265-4222" /></p>
					</div>
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