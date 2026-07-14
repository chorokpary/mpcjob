<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : parnter_skin_olist.asp / 온라인광고 목록 스킨
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-22 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>온라인 광고 관리 - HC 관리자</title>
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
					<h2>온라인 광고 관리</h2>
					<%
					Dim CONF_CATE : CONF_CATE = "5"
					Dim CONF_VIEW : CONF_VIEW = FN_Req("sView","1")
					%>
					<div class="tabArea">
						<span class="tab<%=Fn_SetDefault(CONF_VIEW,"1"," on","")%>"><a href="partner_skin_olist.asp?sView=1">진행중인 사이트</a>
						</span><span class="tab<%=Fn_SetDefault(CONF_VIEW,"0"," on","")%>"><a href="partner_skin_olist.asp?sView=0">종료된 사이트</a></span>
					</div>
					<!--#include file="partner_list.asp"-->
				</div>
			</div>
		</div>
		
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>