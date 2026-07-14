<%
'____________________________________________________________________________________
'
' * Discription : left.asp / 관리자 LNB
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
			<div class="left">
				<div class="inner">
				<% Select Case GNB_STR_MND1%>
					<% Case "01" %>	
					<h2>채용관리</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/recruit/job_skin_list.asp">진행중인 채용공고</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"02")%>><a href="/admin/recruit/job_skin_elist.asp">마감된 채용공고</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"04")%>><a href="/admin/recruit/job_skin_alist.asp">전체 채용공고</a></li>
						<% If Session("ALevel") = "1" Or Session("ALevel") = "2" Then %>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"03")%>><a href="/admin/recruit/job_skin_reg.asp">채용공고 등록</a></li>
						<% End If %>
					</ul>
					<% Case "02" %>	
					<h2>회원관리</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/member/mem_list.asp">회원관리</a></li>
					</ul>
					<% Case "03" %>	
					<h2>프로젝트관리</h2>
					<% Case "04" %>	
					<h2>게시글관리</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/board/notice_list.asp">공지사항</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"02")%>><a href="/admin/board/faq_list.asp">FAQ</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"03")%>><a href="/admin/board/mail_list.asp">Mailling 서비스</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"04")%>><a href="/admin/board/pop_list.asp">팝업관리</a></li>
					</ul>
					<% Case "05" %>	
					<h2>HC통계</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/stats/stats.asp">HC 회원통계</a></li>
					</ul>
					<% Case "06" %>	
					<h2>환경설정</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/siteconf/adm_list.asp">관리자 등록</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"02")%>><a href="/admin/siteconf/conf_main.asp">메인설정</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"03")%>><a href="/admin/siteconf/partner_skin_clist.asp">인력업체관리</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"04")%>><a href="/admin/siteconf/partner_skin_olist.asp">온라인광고관리</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"05")%>><a href="/admin/siteconf/hiretype_list.asp">고용형태관리</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"06")%>><a href="/admin/siteconf/recrpart_list.asp">직종관리</a></li>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"07")%>><a href="/admin/siteconf/ezcareer_list.asp">경력관리</a></li>
					</ul>
					<% Case "07" %>	
					<h2>쪽지함</h2>
					<ul>
						<li <%=FN_AdmGnbOn(GNB_STR_MND2,"01")%>><a href="/admin/message/message_list.asp">쪽지함</a></li>
					</ul>
				<% End Select %>
				</div>
			</div>
