<%
'____________________________________________________________________________________
'
' * Discription : top.asp / 관리자 GNB
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>


		<div class="header">
			<div class="top">
				<div class="inner">
					<p><strong><%=Session("AName")%>님</strong>께서 로그인하셨습니다. [<a href="#btnMyInfo" id="btnMyInfo">정보수정</a>]</p>
					<p class="link"><a href="/admin/logout_proc.asp">로그아웃</a><a href="/admin/">관리자페이지홈</a><a href="javascript:goFrontTest();">홈페이지 바로가기</a></p>
				</div>
			</div>
			<div class="gnb">
				<h1>
					<a href="/admin/"><img src="/admin/images/logo_admin.png" alt="HC Recruit" /></a>
				</h1>
				<p>
					<% If FN_AdmMenuAuth(3) Then %><a href="/admin/recruit/job_skin_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"01")%>>채용관리</a><% End If %>
					<% If FN_AdmMenuAuth(2) Then %><a href="/admin/member/mem_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"02")%>>회원관리</a><% End If %>
					<% If FN_AdmMenuAuth(1) Then %><a href="/admin/project/pjt_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"03")%>>프로젝트관리</a><% End If %>
					<% If FN_AdmMenuAuth(1) Then %><a href="/admin/board/notice_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"04")%>>게시글관리</a><% End If %>
					<% If FN_AdmMenuAuth(1) Then %><a href="/admin/stats/stats.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"05")%>>HC통계</a><% End If %>
					<% If FN_AdmMenuAuth(1) Then %><a href="/admin/siteconf/adm_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"06")%>>환경설정</a><% End If %>
					<% If FN_AdmMenuAuth(1) Then %><a href="/admin/message/message_list.asp" <%=FN_AdmGnbOn(GNB_STR_MND1,"07")%>>쪽지함</a><% End If %>
				</p>
			</div>
		</div>
