<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : recruit_list.asp / 채용공고 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-25 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord
	Dim sRecrPart, sRecrPart2, sHireType
	Dim kData, kFields, i, vNum
	Dim arrList, arrListNum, arrTop, arrTopNum
	
	Call Main()
	
	Sub Main()

		listPage 	= 5
		Page 		= FN_ReqF("Page","1")
		listRecord 	= FN_ReqF("ListSize",15)	'목록갯수
		
		sKey 	= FN_ReqF("sKey","ALL")
		sWord 	= FN_InputVal(FN_ReqF("sWord",""))

		sRecrPart 	= FN_ReqF("sRecrPart","")
		sRecrPart2 	= FN_ReqF("sRecrPart2","")
		sHireType 	= FN_ReqF("sHireType","")
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If

		If IsNumeric( listRecord ) Then
			listRecord = CInt( listRecord )
		Else
			listRecord = 15
		End If
		
		Call getTopList()
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@sPart").Value 	= GLOBAL_SEARCH_PART
			.Parameters("@sAddrSido").Value = GLOBAL_SEARCH_SIDO
			.Parameters("@sAddrGugun").Value = GLOBAL_SEARCH_GUGUN
			.Parameters("@sRecrPart").Value = sRecrPart
			.Parameters("@sRecrPart2").Value = sRecrPart2
			.Parameters("@sHireType").Value = sHireType

			.Parameters("@sViewFlag").Value = ""	' 전체글
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrListNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","PrjName","LogoPath","AppEDate")
			arrList 	= Rs.GetRows(,,kFields)
			arrListNum 	= UBound(arrList, 2)
			kData 		= True

			RecordCount = arrList(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

	
	Sub getTopList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FRT"
			.Parameters("@pno").Value 		= 1
			.Parameters("@per_page").Value 	= 4

			.Parameters("@sViewFlag").Value = "TOP"
			
			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			arrTopNum 	= -1
		Else
			' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
			
			kFields 	= Array("TotalCnt","RecrSeq","Title","PrjName","LogoPath","AppEDate")
			arrTop 	= Rs.GetRows(,,kFields)
			arrTopNum 	= UBound(arrTop, 2)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>채용정보 - HANKOOK CORPORATION</title>
	<!--#include virtual="/common/include/head.asp"-->
	<script type="text/javascript">
		var RecrPart2 = "<%=sRecrPart2%>";
		function goView(seq){
			goViewCommon(seq,"recruit_view.asp");
		}
		
		function doClearSearch(){
			$("#Page").val("1");
			$("#sPart").val("");
			$("#sAddrSido").val("");
			$("#sAddrGugun").val("");
		}

		function getRecrPart2(){
			if ($("#sRecrPart").val()) {
				var url = "/common/dev/ajax/make_select_cd_2depth.asp";
				var param = "Flag=RecrPart&Parent=" + $("#sRecrPart").val() + "&TargetID=sRecrPart2&Def=" + RecrPart2 + "&Css=w180";
				
				RecrPart2 = "";
				
				$.ajax({
					type: "POST"
					, url: url
					, data: param
					, success: function(content){ $("#divRecrPart2").html(content); }
					, error: function(xhr, status, error) {alert("[Rec] " + error);}
				});
			}else{
				$("#divRecrPart2").html("");
			}
		}
		
		function setListSize(num){
			$("#Page").val("1");
			$("#ListSize").val(num);
			$("#frmPage").submit();
			
		}
		
		$(function() {
		    $("#sWord").keyup(doClearSearch);
		    $("#sRecrPart").change(function(){ getRecrPart2(); doClearSearch();});
		    $("#sRecrPart2").change(doClearSearch);
		    $("#sHireType").change(doClearSearch);
		    
		    $(window).load(getRecrPart2);
		});
		
		
	</script>
</head>
<body>
	<div id="wrap" class="recruitList">
		<!--#include virtual="/common/include/top.asp"-->
		<div id="container">
			<div class="contents">
				<div class="clearfix bread">
					<div class="fl_l">
						<h2><img src="/images/pages/tit_recruit.jpg" alt="채용정보" /></h2>
					</div>
					<div class="fl_r">
						<p><%=GNB_STR_LOC%></p>
					</div>
				</div>

				<form name="frmPage" id="frmPage" method="get">
				<input type="hidden" id="Seq" name="Seq">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				<input type="hidden" id="ListSize" name="ListSize" value="<%=listRecord%>">
				<input type="hidden" id="sPart" name="sPart" value="<%=GLOBAL_SEARCH_PART%>">
				<input type="hidden" id="sAddrSido" name="sAddrSido" value="<%=GLOBAL_SEARCH_SIDO%>">
				<input type="hidden" id="sAddrGugun" name="sAddrGugun" value="<%=GLOBAL_SEARCH_GUGUN%>">
				<div class="boxArea search">
					<p>
						<label for="sWord">통합검색</label> <input type="text" name="sWord" id="sWord" title="통합검색" value="<%=sWord%>" class="w360" />
					</p>
					<p>
						<label for="sRecrPart">상세검색</label>
						<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "RecrPart", "sRecrPart", "직종별", "", sRecrPart, "등록된 분야가 없습니다.", " class=""w180"" ", False)%>
						<span id="divRecrPart2" style="margin:0; padding:0;"><!-- 직종코드 2Depth //--></span>
						<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_RECRUIT", "HireType", "sHireType", "고용형태", "", sHireType, "등록된 고용형태가 없습니다.", " class=""w180"" style=""margin-left:5px;"" ", False)%>
					</p>
					<input type="image" src="/images/pages/btn_search.jpg" alt="검색하기" class="btn" />
				</div>
				</form>
				
				<% If arrTopNum > -1 Then %>
				<div class="special">
					<h3><img src="/images/pages/stit_recruit_01.jpg" alt="스페셜 채용정보" /></h3>
					<ul class="pjtList">
						<% 		
								' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
								For i = 0 To arrTopNum 
						%>
						<li>
							<a href="javascript:goView(<%=arrTop(1,i)%>);">
								<div class="imgs">
									<img alt="<%=arrTop(3,i)%>" src="<%=Fn_SetDefault(arrTop(4,i),"","/images/common/no_company.jpg",arrTop(4,i))%>">
									<div class="bg"></div>
								</div>
							</a>
							<div class="info">
								<h4><a href="javascript:goView(<%=arrTop(1,i)%>);"><%=arrTop(3,i)%></a></h4>
								<p class="pjt"><%=arrTop(2,i)%></p>
								<p class="due">서류마감 : <%=FN_SetDateTimeFormat(arrTop(5,i),"MM/DD")%></p>
							</div>
						</li>

						<%
								Next 
						%>
					</ul>
				</div>
				<% End If %>
				
				<div class="normal">
					<h3><img src="/images/pages/stit_recruit_02.jpg" alt="주요 채용정보" /></h3>
					<div class="listsize">
						<a href="javascript:setListSize(0);"<%=Fn_SetDefault(listRecord,"0"," class=""on"" ","")%>>모두보기</a> | <a href="javascript:setListSize(15);"<%=Fn_SetDefault(listRecord,"15"," class=""on"" ","")%>>15개씩 보기</a>
					</div>
					<ul class="pjtList">
						<% If Not kData Then %>
						<li class="nolist">현재 진행중인 공고가 없습니다.</li>
						<% Else %>
						<% 		
								' TotalCnt(0), RecrSeq(1), Title(2), PrjName(3), LogoPath(4), AppEDate(5)
								vNum = RecordCount - ((Page-1) * listRecord) 
								
								For i = 0 To arrListNum 
						%>
						<li>
							<a href="javascript:goView(<%=arrList(1,i)%>);">
								<div class="imgs">
									<img alt="<%=arrList(3,i)%>" src="<%=Fn_SetDefault(arrList(4,i),"","/images/common/no_company.jpg",arrList(4,i))%>">
									<div class="bg"></div>
								</div>
							</a>
							<div class="info">
								<h4><a href="javascript:goView(<%=arrList(1,i)%>);"><%=arrList(3,i)%></a></h4>
								<p class="pjt"><%=arrList(2,i)%></p>
								<p class="due">서류마감 : <%=FN_SetDateTimeFormat(arrList(5,i),"MM/DD")%></p>
							</div>
						</li>
						<%			vNum = vNum - 1
								Next 
						%>
						<% End If %>
					</ul>
					<div class="paging">
						<% Call SB_PagingWrite(listPage, RecordCount, PageCount) %>
					</div>
				</div>
				
			</div>
			<!--#include virtual="/common/include/left.asp"-->
		</div>
		<!--#include virtual="/common/include/footer.asp"-->
	</div>
	<!--#include virtual="/common/include/popup.asp"-->
</body>
</html>