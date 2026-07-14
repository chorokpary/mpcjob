<%
'____________________________________________________________________________________
'
' * Discription : left.asp / Front LNB
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<div class="leftM">
<% Select Case GNB_STR_MND1%>
	<% Case "03" %>	
	<div class="lnbType3">
		<h2><img src="/images/pages/tit_lnbCorp.png" alt="About HC" /></h2>
		<ul>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"01")%>><a href="/corp/philosophy.asp">회사소개</a></li>
			<!-- <li <%=FN_FrontGnbOn(GNB_STR_MND2,"02")%>><a href="/corp/center.asp">센터찾기</a></li> -->
		</ul>
	</div>
	<% Case "04" %>	
	<div class="lnbType1">
		<h2><img src="/images/pages/tit_lnbInform.png" alt="HC알리미" /></h2>
		<ul>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"01")%>><a href="/inform/notice_list.asp">공지사항</a></li>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"02")%>><a href="/inform/faq_list.asp">FAQ</a></li>
		</ul>
	</div>
	<% Case "05" %>	
	<div class="lnbType2">
		<h2><img src="/images/pages/tit_lnbMypage.png" alt="mypage" /></h2>
		<ul>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"01")%>><a href="/mypage/resume.asp">이력서관리</a></li>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"02")%>><a href="/mypage/apply.asp">입사지원관리</a></li>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"03")%>><a href="/mypage/matchArea.asp">인근지역공고</a></li>
			<li <%=FN_FrontGnbOn(GNB_STR_MND2,"04")%>><a href="/mypage/memSecession.asp">회원탈퇴</a></li>
		</ul>
	</div>
	<% Case "07" %>	
<% End Select %>
	<form name="frmSearch_Part" id="frmSearch_Part" method="post" action="/recruit/recruit_list.asp">
	<input type="hidden" name="Search_Part" id="Search_Part" value="<%=FN_Req("GLOBAL_SEARCH_PART","")%>">
	<div class="cate clearfix">
		<a href="javascript:setSearch_Part('공공기관');"><img src="/images/common/link_sub_01_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"공공기관")%>.png" alt="공공기관" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"공공기관")%>/></a>
		<a href="javascript:setSearch_Part('통신분야');"><img src="/images/common/link_sub_02_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"통신분야")%>.png" alt="통신분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"통신분야")%>/></a>
		<a href="javascript:setSearch_Part('금융분야');"><img src="/images/common/link_sub_03_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"금융분야")%>.png" alt="금융분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"금융분야")%> /></a>
		<a href="javascript:setSearch_Part('IT분야');"><img src="/images/common/link_sub_04_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"IT분야")%>.png" alt="IT분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"IT분야")%> /></a>
		<a href="javascript:setSearch_Part('쇼핑분야');"><img src="/images/common/link_sub_05_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"쇼핑분야")%>.png" alt="쇼핑분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"쇼핑분야")%> /></a>
		<a href="javascript:setSearch_Part('교육분야');"><img src="/images/common/link_sub_06_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"교육분야")%>.png" alt="교육분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"교육분야")%> /></a>
		<a href="javascript:setSearch_Part('건설분야');"><img src="/images/common/link_sub_07_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"건설분야")%>.png" alt="건설분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"건설분야")%> /></a>
		<a href="javascript:setSearch_Part('기타분야');"><img src="/images/common/link_sub_08_<%=FN_FrontImgOn(GLOBAL_SEARCH_PART,"기타분야")%>.png" alt="기타분야" <%=FN_FrontGnbOn(GLOBAL_SEARCH_PART,"기타분야")%> /></a>
	</div>
	</form>
	<form name="frmLnbAreaSearch" id="frmLnbAreaSearch" method="post" action="/recruit/recruit_list.asp">
	<div class="search_loc">
		<h2><img src="/images/common/tit_search_loc.png" alt="찾으시는 지역을 검색하세요" /></h2>
		<p><img src="/images/common/txt_search_loc.png" alt="검색을 통해 각 권역별 채용정보를 알려드립니다." /></p>
		<%
		Dim Search_SidoSql : Search_SidoSql = "SELECT SIDO AS CODE, SIDO AS NAME FROM TBL_ZIPCODE GROUP BY SIDO ORDER BY SIDO" 
		response.write " ": Call SB_MakeHTMLSelect(objDBCon, Search_SidoSql, "Search_AddrSido", "지역", "", GLOBAL_SEARCH_SIDO, "", " class=""selRelSite"" ", False)
		%>
		<div id="divSearchGugun">
		<%
		Dim Search_GugunSql : Search_GugunSql = "SELECT GUGUN AS CODE, GUGUN AS NAME FROM TBL_ZIPCODE WHERE SIDO = '" &  GLOBAL_SEARCH_SIDO & "' GROUP BY GUGUN ORDER BY GUGUN" 
		response.write " ": Call SB_MakeHTMLSelect(objDBCon, Search_GugunSql, "Search_AddrGugun", "지역중분류", "",  GLOBAL_SEARCH_GUGUN, "", " class=""selRelSite"" ", False)
		%>
		</div>
		<input type="image" src="/images/common/search_loc_sub.png" alt="지역검색하기" />
	</div>
	</form>
	<div class="csCall">
		<h2><img src="/images/common/tit_sub_cs.png" alt="한국코퍼레이션 채용문의 02-3432-6100" /></h2>
		<p><img src="/images/common/txt_sub_cs.png" alt="h.p:010-9265-4222 / mail:hr@besthc.co.kr / 평일:AM 9:00~PM6:00 / 점심:PM12:00~PM1:00 / 주말 및 공휴일 휴무" /></p>
	</div>
	<!-- 20190412 포엠 카카오톡 추가 -->
	<div class="csCall kakao">
	    <img src="/images/common/kakao_tit_cs.png" alt="카카오톡 채팅상담" />
	    <a href="http://pf.kakao.com/_zxbPrj" class="kakao_link" target="_blank"><img src="/images/common/kakao_cs.png" alt="한국코퍼레이션 카카오톡" /></a>
	</div>
	<!-- //20190412 포엠 카카오톡 추가 -->
</div>