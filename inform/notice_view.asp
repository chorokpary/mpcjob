<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : notice_view.asp / 공지사항 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-23 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim sKey, sWord
	Dim Subject, Attach1_Path, Attach1_Name, Attach2_Path, Attach2_Name, Contents
	Dim RegDate, RegName, Hit
	Dim nextSeq, nextSubject, nextDate, prevSeq, prevSubject, prevDate
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_Req("sWord","")
		
		If Seq <> "0" Then
			Call getData
			Flag = "MOD"
		Else
			Flag = "ADD"
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspBbs_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "INFO"
			.Parameters("@BbsSeq").Value 	= Seq
			.Parameters("@BbsKey").Value 	= "notice"
			.Parameters("@UseCnt").Value 	= 1
			.Parameters("@UsePN").Value 	= 1
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				Subject = FN_InputVal(Rs("Subject"))
				Attach1_Path = Rs("Attach1_Path")
				Attach1_Name = Rs("Attach1_Name")
				Attach2_Path = Rs("Attach2_Path")
				Attach2_Name = Rs("Attach2_Name")
				Contents = Rs("Contents")
				RegDate = Rs("RegDate")
				RegName = Rs("RegName")
				Hit = Rs("Hit")
				'FN_setBlank
			End If

			Set Rs = Rs.NextRecordSet
			
			If Not Rs.Eof Then
				prevSeq = Rs("BbsSeq")
				prevSubject = Rs("Subject")
				prevDate = FN_SetDateTimeFormat(Rs("RegDate"),"YYYY-MM-DD")
			End If
	
			Set Rs = Rs.NextRecordSet
			
			If Not Rs.Eof Then
				nextSeq = Rs("BbsSeq")
				nextSubject = Rs("Subject")
				nextDate = FN_SetDateTimeFormat(Rs("RegDate"),"YYYY-MM-DD")
			End If
		
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>HC알리미 - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function goView(seq){
			goViewCommon(seq, "notice_view.asp");
		}

		function goList(){
			goListCommon("notice_list.asp");
		}
	</script>
</head>
<body>
	<div id="wrap">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_notice.jpg" alt="공지사항" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>
				<div class="board_view">
					<p class="thisTit"><%=Subject%></p>
					<table>
						<caption>게시글 정보</caption>
						<colgroup>
							<col class="w80" />
							<col />
							<col class="w80" />
							<col />
							<col class="w80" />
							<col />
						</colgroup>
						<tbody>
							<tr>
								<th><label for="">작성자</label></th>
								<td><%=RegName%></td>
								<th><label for="">등록일</label></th>
								<td><%=FN_SetDateTimeFormat(RegDate,"YYYY/MM/DD")%></td>
								<th><label for="">조회수</label></th>
								<td><%=Hit%></td>
							</tr>
							<% If Attach1_Path <> "" Or Attach2_Path <> "" Then%>
							<tr class="attach">
								<th><label for="">첨부파일</label></th>
								<td colspan="5">
									<% If Attach1_Path <> "" Then %><p><a href="javascript:goDown('<%=Seq%>','BBS','Attach1_Path','<%=Attach1_Name%>');"><%=Attach1_Name%></a></p><% End If %>
									<% If Attach2_Path <> "" Then %><p><a href="javascript:goDown('<%=Seq%>','BBS','Attach1_Path','<%=Attach2_Name%>');"><%=Attach2_Name%></a></p><% End If %>
								</td>
							</tr>
							<% End If %>
						</tbody>
					</table>
					<div class="inner"><%=Contents%></div>
				</div>
				
				<div class="viewBottom">
					<ul>
						<li class="prevTit">이전글</li>
						<li>
							<% If prevSeq > 0 Then %>
							<a href="javascript:goView(<%=prevSeq%>);"><%=prevSubject%></a>
							<% Else %>
							이전글이 없습니다.
							<% End If %>
						</li>
					</ul>
					<ul>
						<li class="nextTit">다음글</li>
						<li>
							<% If nextSeq > 0 Then %>
							<a href="javascript:goView(<%=nextSeq%>);"><%=nextSubject%></a>
							<% Else %>
							다음글이 없습니다.
							<% End If %>
						</li>
					</ul>
				</div>
				<p class="rightBtn"><a href="javascript:goList();" id="btnList"><img src="/images/common/btn_list.gif" alt="목록" /></a></p>
				
				<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				<input type="hidden" id="Seq" name="Seq">
				<input type="hidden" id="sKey" name="sKey" value="<%=sKey%>">
				<input type="hidden" id="sWord" name="sWord" value="<%=sWord%>">
				</form>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>