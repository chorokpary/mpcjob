<div id="header">
			<div class="top">
				<h1><a href="/"><img src="/images/common/logo.png" alt="HANKOOK CORPORATION" /></a></h1>
				<div class="member">
					<a href="/"><img src="/images/common/top_home.png" alt="홈" />
					<% If Session("USeq") = "" Then %>
					</a><a href="/membership/login.asp"><img src="/images/common/top_login.png" alt="로그인" />
					</a><a href="/membership/easySignUp_n.asp"><img src="/images/common/top_join.png" alt="회원가입" /></a>
					<a href="/recruit/recruit_view_t.asp?Seq=28329"><img src="/images/common/top_recommend.png" alt="직원추천 입사지원" /></a>
					<% Else %>
					</a><a href="/membership/logout.asp"><img src="/images/common/top_logout.png" alt="로그아웃" />
					</a><a href="/mypage/resume.asp"><img src="/images/common/top_privacy.png" alt="개인정보수정" /></a>
					<a href="/recruit/recruit_view_t.asp?Seq=28329"><img src="/images/common/top_recommend.png" alt="직원추천 입사지원" /></a>
					<% End If %>
				</div>
			</div>
			<div class="gnb clearfix">
				<ul>
					<li <%=FN_FrontGnbOn(GNB_STR_MND1,"01")%>><a href="/recruit/recruit_list.asp"><img src="/images/common/gnb_recruit.png" alt="채용정보" /></a></li>
					<li <%=FN_FrontGnbOn(GNB_STR_MND1,"02")%>><a href="/family/family.asp"><img src="/images/common/gnb_family.png" alt="HC Family"/></a></li>
					<li <%=FN_FrontGnbOn(GNB_STR_MND1,"03")%>><a href="/corp/philosophy.asp"><img src="/images/common/gnb_mpc.png" alt="About HC" /></a></li>
					<li <%=FN_FrontGnbOn(GNB_STR_MND1,"04")%>><a href="/inform/notice_list.asp"><img src="/images/common/gnb_notification.png" alt="HC 알리미" /></a></li>
					<li <%=FN_FrontGnbOn(GNB_STR_MND1,"05")%>><a href="/mypage/resume.asp"><img src="/images/common/gnb_mypage.png" alt="MY PAGE" /></a></li>
				</ul>
				<div class="menuAll">
					<a href="javascript:goFullMenu();"><img src="/images/common/gnb_menuAll.png" alt="전체메뉴보기" /></a>
				</div>
			</div>
			<div class="links">
				<p><%=GNB_STR_GREETING%></p>
				<p class="download">
					이력서 양식 다운로드
					<a href="javascript:goDown('0','CONF','ConfFormDoc','');"><img src="/images/common/ico_top_word.png" alt="워드" /></a>
					<a href="javascript:goDown('0','CONF','ConfFormXls','');"><img src="/images/common/ico_top_excel.png" alt="엑셀" /></a>
					<a href="javascript:goDown('0','CONF','ConfFormHwp','');"><img src="/images/common/ico_top_hwp.png" alt="한글" /></a>
				</p>
			</div>
</div>