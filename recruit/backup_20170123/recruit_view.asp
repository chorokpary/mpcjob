<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : recruit_view.asp / 채용공고 상세보기
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-11 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page, PageFlag
	Dim sKey, sWord
	Dim sRecrPart, sRecrPart2, sHireType
	Dim PrjName, Title, IsMain, IsTop, ChargeTask, TaskType
	Dim Part, RecrPart, RecrPart2, Career, Pay, HireType, LocalSido, LocalGugun
	Dim AppBDate, AppEDate, MapPath, LimitEdu, LimitCareer, LimitBAge, LimitEAge, LimitGender
	Dim ChargeName, ChargePhone, ChargeEmail, Contents, IsView

	Dim i, yy, mm
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		Page = FN_Req("Page","1")
		PageFlag = FN_Req("PageFlag","")

		sKey 	= FN_Req("sKey","ALL")
		sWord 	= FN_InputVal(FN_Req("sWord",""))

		sRecrPart 	= FN_Req("sRecrPart","")
		sRecrPart2 	= FN_Req("sRecrPart2","")
		sHireType 	= FN_Req("sHireType","")
		
		If sWord <> "" Then
			GLOBAL_SEARCH_PART = ""
			GLOBAL_SEARCH_SIDO = ""
			GLOBAL_SEARCH_GUGUN = ""
		End If 
		
		If PageFlag = "ADM" Or Seq = "0"  Then
			Call getPreview
		Else
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			.Parameters("@UseHit").Value = 1
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PrjName 	= Rs("PrjName")
				Title 		= FN_InputVal(Rs("Title"))
				IsMain 		= Rs("IsMain")
				IsTop 		= Rs("IsTop")
				ChargeTask 	= FN_InputVal(Rs("ChargeTask"))
				TaskType 	= FN_InputVal(Rs("TaskType"))
				Part 		= Rs("Part")
				RecrPart 	= Rs("RecrPart")
				RecrPart2 	= Rs("RecrPart2")
				Career 		= Rs("Career")
				Pay 		= FN_InputVal(Rs("Pay"))
				HireType 	= Rs("HireType")
				LocalSido 	= Rs("LocalSido")
				LocalGugun 	= Rs("LocalGugun")
				AppBDate 	= FN_SetDateTimeFormat(Rs("AppBDate"),"YYYY-MM-DD") 
				AppEDate 	= FN_SetDateTimeFormat(Rs("AppEDate"),"YYYY-MM-DD") 
				MapPath 	= Rs("MapPath")
				ChargeName 	= Rs("AdmName")
				ChargePhone = FN_InputVal(Rs("ChargePhone"))
				ChargeEmail = FN_InputVal(Rs("ChargeEmail"))
				Contents 	= Rs("Contents")
				IsView		= Rs("IsView")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
	Sub getPreview()
		PrjName 	= FN_Req("PrjName","")
		Title 		= FN_Req("Title","")
		ChargeTask 	= FN_Req("ChargeTask","")
		TaskType 	= FN_Req("TaskType","")
		Part 		= FN_Req("Part","")
		RecrPart 	= FN_Req("RecrPart","")
		RecrPart2 	= FN_Req("RecrPart2","")
		Career 		= FN_Req("Career","")
		Pay 		= FN_Req("Pay","")
		HireType 	= FN_Req("HireType","")
		LocalSido 	= FN_Req("LocalSido","")
		LocalGugun 	= FN_Req("LocalGugun","")
		AppBDate 	= FN_SetDateTimeFormat(FN_Req("AppBDate",""),"YYYY-MM-DD") 
		AppEDate 	= FN_SetDateTimeFormat(FN_Req("AppEDate",""),"YYYY-MM-DD") 

		ChargeName 	= FN_Req("ChargeName","")
		ChargePhone = FN_Req("ChargePhone","")
		ChargeEmail = FN_Req("ChargeEmail","")
		Contents 	= FN_Req("Contents","")
		
		MapPath 	= FN_Req("TmpImg","") 
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>채용정보 - MPC JOB</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		function doApply(){
			<% If PageFlag = "ADM" Or Seq = "0" Then%>
			alert("미리보기에서는 지원하지 않는 기능입니다.")
			<% Else %>
				<% If Session("USeq") = "" Then %>
				//로그인 요청창 출력
				openLayer("/membership/loginRequest.asp","RecrSeq=<%=Seq%>", "w668");
				<% Else %>
				//입사지원
				openLayer("/recruit/apply_proc.asp","RecrSeq=<%=Seq%>", "w668");
				<% End If %>
			<% End If%>
		}

		function doQNA(){
			<% If PageFlag = "ADM" Or Seq = "0" Then%>
			alert("미리보기에서는 지원하지 않는 기능입니다.")
			<% Else %>
				<% If Session("USeq") = "" Then %>
				// 로그인
				$("#frmPage").attr({action:"/membership/login.asp", method:"get"}).submit();
				<% Else %>
				// 문의하기
				openLayer("inquiryPop.asp","RecrSeq=<%=Seq%>", "w668");
				<% End If %>
			<% End If %>
		}
	
		$(function() {
			$("#btnApply").click(doApply);
			$("#btnQNA").click(doQNA);
		});
		
		$(window).load(function(){
			<% If (PageFlag = "ADM" Or Seq = "0") And MapPath = "" Then %>
			// 미리보기 > 로컬이미지
			var obj = opener.document.getElementById("MapPath");
			var target = document.getElementById("divMap");
			fileUploadPreview(obj,target);
			<% End If %>
			
			$("#frmContents").submit();
		})
	</script>
</head>
<body>
	<div id="wrap" class="recruitView">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_recruit.jpg" alt="채용정보" /></h2>
					</div>
					<div class="fl_r">
						<p>HOME : 채용정보</p>
					</div>
				</div>
				
				<form name="frmPage" id="frmPage" method="post" action="recruit_list.asp">
				<div class="boxArea height68">
					<p>
						<label for="sWord">통합검색</label> <input type="text" name="sWord" id="sWord" title="통합검색" value="<%=sWord%>" class="w360" />
					</p>
					<input type="image" src="/images/pages/btn_search.jpg" alt="검색하기" class="btn" />
				</div>
				</form>
				
				<div class="guideline">
					<h3><img src="/images/pages/stit_recruit_03.jpg" alt="모집요강" /></h3>
					<h4><%=Title%></h4>
					<table>
						<caption>모집요강</caption>
						<colgroup>
							<col class="w80" />
							<col />
							<col class="w80" />
							<col />
						</colgroup>
						<tbody>
							<tr>
								<th><label for="">회사명</label></th>
								<td colspan="3"><%=PrjName%></td>
							</tr>
							<tr>
								<th><label for="">분야별</label></th>
								<td><%=Part%></td>
								<th><label for="">모집기간</label></th>
								<td>
									<%=AppBDate%>
									<% If AppBDate <> "" And AppEDate <> "" Then Response.Write " ~ "%>
									<%=AppEDate%>
								</td>
							</tr>
							<tr>
								<th><label for="">업무형태</label></th>
								<td><%=TaskType%></td>
								<th><label for="">직종별</label></th>
								<td><%=RecrPart%><%=Fn_SetDefault(RecrPart2,"",""," > " & RecrPart2)%></td>
							</tr>
							<tr>
								<th><label for="">근무지역</label></th>
								<td><%=LocalSido%>&nbsp;<%=LocalGugun%></td>
								<th><label for="">고용형태</label></th>
								<td><%=HireType%></td>
							</tr>
						</tbody>
					</table>
					<div class="cont">
						<iframe id="ifrmContents" name="ifrmContents" width="100%" frameborder="0" scrolling="no"></iframe>
						<form name="frmContents" id="frmContents" method="post" action="recruit_view_iframe.asp" target="ifrmContents">
						<textarea id="txtContents" name="txtContents" style="display:none;"><%=Contents%></textarea>
						</form>
					</div>
				</div>
				<div class="guideline">
					<h3><img src="/images/pages/stit_recruit_04.jpg" alt="문의 및 담당자 정보" /></h3>
					<table>
						<caption>문의 및 담당자정보</caption>
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
								<th><label for="">담당자</label></th>
								<td><%=ChargeName%></td>
								<th><label for="">이메일</label></th>
								<td><%=ChargeEmail%></td>
								<th><label for="">연락처</label></th>
								<td><%=ChargePhone%></td>
							</tr>
						</tbody>
					</table>
					<div class="cont" id="divMap"><%=Fn_SetDefault(MapPath,"","","<img src=""" & MapPath & """>")%></div>
				</div>
				<div class="btn_apply">
					<a href="javascript:;" id="btnApply"><img src="/images/pages/btn_recruit.jpg" alt="온라인입사지원" /></a><a 
					href="javascript:;" id="btnQNA"><img src="/images/pages/btn_inquiry.jpg" alt="문의하기" /></a>
				</div>
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>