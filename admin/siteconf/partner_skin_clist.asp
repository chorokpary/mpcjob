<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : parnter_skin_clist.asp / 인력업체 목록 스킨
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>인력업체관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea">
					<h2>인력업체관리</h2>
					<%
					Dim CONF_CATE : CONF_CATE = "6"
					Dim CONF_VIEW : CONF_VIEW = FN_Req("sView","1")
					%>
					<div class="tabArea">
						<span class="tab<%=Fn_SetDefault(CONF_VIEW,"1"," on","")%>"><a href="partner_skin_clist.asp?sView=1">진행중인 인력업체</a>
						</span><span class="tab<%=Fn_SetDefault(CONF_VIEW,"0"," on","")%>"><a href="partner_skin_clist.asp?sView=0">종료된 인력업체</a></span>
					</div>
					
					<!--#include file="partner_list.asp"-->
				</div>
			</div>
		</div>
		
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>