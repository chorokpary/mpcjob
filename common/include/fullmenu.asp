<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
	<div class="menuInner">
		<h2><img src="/images/common/tit_menu.png" alt="HANKOOK CORPORATION 전체메뉴 보기" /></h2>
		<ul class="menuTop">
			<% If Session("USeq") = "" Then %>
			<li><a href="/membership/login.asp"><img src="/images/common/menu_login.png" alt="로그인" /></a></li>
			<% Else %>
			<li><a href="/membership/logout.asp"><img src="/images/common/menu_logout.png" alt="로그아웃" /></a></li>
			<li><a href="/mypage/resume.asp"><img src="/images/common/menu_ingoMod.png" alt="개인정보수정" /></a></li>
			<% End If %>
		</ul>
		<div class="munuList">
			<ul>
				<li><h3><a href="/recruit/recruit_list.asp"><img src="/images/common/menu_recruit.png" alt="채용정보" /></a></h3>
					<div>각종 채용정보를 경험하실 수 있습니다.</div>
				</li>
				<li><h3><a href="/family/family.asp"><img src="/images/common/menu_family.png" alt="HC Family" /></a></h3>
					<div>한국코퍼레이션 한가족을 만나보세요.</div>
				</li>
				<li><h3><a href="/corp/philosophy.asp"><img src="/images/common/menu_corp.png" alt="About HC" /></a></h3>
					<div>
						<h4><a href="/corp/philosophy.asp">회사소개</a></h4>
						<p><a href="/corp/philosophy.asp">경영철학</a> 
						/ <a href="/corp/people.asp">인재상</a> 
						/ <a href="/corp/benefit.asp">급여 및 복리후생</a> 
						/ <a href="/corp/eduSystem.asp">교육제도</a></p>
					</div>
				</li>
				<li><h3><a href="/inform/notice_list.asp"><img src="/images/common/menu_inform.png" alt="HC 알리미" /></a></h3>
					<div>
						<h4><a href="/inform/notice_list.asp">공지사항</a></h4><h4><a href="/inform/faq_list.asp">FAQ</a></h4>
					</div>
				</li>
				<li><h3><a href="/mypage/resume.asp"><img src="/images/common/menu_myPage.png" alt="MY PAGE" /></a></h3>
					<div>
						<h4><a href="/mypage/resume.asp">이력서관리</a></h4>
						<h4><a href="/mypage/apply.asp">입사지원관리</a></h4>
						<h4><a href="/mypage/matchArea.asp">인근지역공고</a></h4>
						<h4><a href="/mypage/memSecession.asp">회원탈퇴</a></h4>
					</div>
				</li>
			</ul>
		</div>
		<p class="btnClose"><a href="#" id="btnClose"><img src="/images/common/menu_close.png" alt="닫기" /></a></p>
	</div>
