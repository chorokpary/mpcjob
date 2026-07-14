<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pjt_list.asp / 프로젝트 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, sView, sFlag
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	Dim strTabFlag

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 30	'목록갯수
		Page 		= FN_Req("Page","1")

		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		sView 	= FN_Req("sView","1")
		sFlag 	= FN_Req("sFlag","")
		
		If sView = "1" And sFlag = "" Then	sFlag = "도급" 
		
		If sView = "1" And sFlag = "도급" Then
			strTabFlag = 1
		ElseIf sView = "1" And sFlag = "파견" Then
			strTabFlag = 2
		ElseIf sView = "0" Then
			strTabFlag = 3
		Else
			Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		End If
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		
	End Sub
	
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspProject_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@sFlag").Value 	= sFlag
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			.Parameters("@sView").Value 	= sView
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), PrjSeq(1), PrjName(2), LogoPath(3), Flag(4)
			kFields 	= Array("TotalCnt","PrjSeq","PrjName","LogoPath","Flag")
			arrData 	= Rs.GetRows(,,kFields)
			arrDataNum 	= UBound(arrData, 2)
			kData 		= True
			
			RecordCount = arrData(0,0)
			PageCount = FN_PageCount(RecordCount,listRecord)
		End If
		
		Rs.Close
		Set Rs = Nothing

	End Sub

%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>프로젝트관리 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
		function chkDel(){
			if (cntChecked("#frmList")==0){
				alert("선택된 데이터가 없습니다.");
			}else{
				if(confirm("선택한 데이터를 <%=Fn_SetDefault(strTabFlag,"3","복구","삭제")%>하시겠습니까?")){
					$("#Flag").val("STA");
					$("#frmList").attr({action:"pjt_proc.asp", method:"post"}).submit();
				}
			}
		}
		
		function openForm(seq){
			openWin("pjt_form.asp?Seq="+seq,"popPartner",0,0,930,650,0);
		}
		
		
		$(function() {
		    $("#btnChkAll").click(function(){chekAllBtn("#frmList","#btnChkAll")});
		    $("#btnStatus").click(chkDel);
		    $("#btnReg").click(function(){ openForm(0);});
		});
	</script>
</head>

<body>
	<div id="wrap">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location"><%=GNB_STR_LOC%></div>
				<div id="contArea" class="pjt_list">
					<h2>프로젝트관리</h2>
					
					
					<div class="tabArea">
						<span class="tab<%=Fn_SetDefault(strTabFlag,1," on","")%>"><a href="pjt_list.asp?sFlag=도급&sView=1">진행중인 프로젝트</a>
						</span><span class="tab<%=Fn_SetDefault(strTabFlag,2," on","")%>"><a href="pjt_list.asp?sFlag=파견&sView=1">진행중인 파견 프로젝트</a>
						</span><span class="tab<%=Fn_SetDefault(strTabFlag,3," on","")%>"><a href="pjt_list.asp?sView=0">제외된 프로젝트</a></span>
					</div>
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_g"><a href="#btnChkAll" id="btnChkAll">전체선택</a>
							</span><span class="btns btn_g"><a href="#btnStatus" id="btnStatus"><%=Fn_SetDefault(strTabFlag,"3","복구","삭제")%></a>
							</span><span class="btns btn_b"><a href="#btnReg" id="btnReg">프로젝트등록</a></span>
						</div>
						<div class="fl_r">
						<form name="frmPage" id="frmPage" method="post">
							<input type="hidden" id="Page" name="Page" value="<%=Page%>">
							<input type="hidden" id="sKey" name="sKey" value="NAME">
							<input type="hidden" id="sView" name="sFlag" value="<%=sFlag%>">
							<input type="hidden" id="sView" name="sView" value="<%=sView%>">
							<input type="text" id="sWord" name="sWord" value="<%=sWord%>" class="w190" /><span class="btns btn_b"><button type="submit">검색</button></span>
						</form>
						</div>
					</div>
					<div class="pjtListArea">
						<div class="inner">
							<form name="frmList" id="frmList">
							<input type="hidden" id="Flag" name="Flag">
							<input type="hidden" id="sFlag" name="sFlag" value="<%=sFlag%>">
							<input type="hidden" id="sView" name="sView" value="<%=sView%>">
							<ul>
								<% If Not kData Then %>
								<li class="notList">등록된 프로젝트가 없습니다.</li>
								<% Else %>
								<% 		
										' TotalCnt(0), PrjSeq(1), PrjName(2), LogoPath(3), Flag(4)
										vNum = RecordCount - ((Page-1) * listRecord) 
										For i = 0 To arrDataNum 
								%>
								<li>
									<a href="javascript:openForm('<%=arrData(1,i)%>')">
									<div class="imgs">
										<img src="<%=Fn_SetDefault(arrData(3,i),"","/admin/images/no_company.jpg",arrData(3,i))%>" alt="" />
										<div class="bg"></div>
										<% If arrData(4,i) = "파견" Then %><img src="/admin/images/ico_dispatch.png" alt="파견" class="send" /><% End If %>
									</div>
									</a>
									
									<p  class="info">
										<input type="checkbox" name="Seq" id="Seq_<%=arrData(1,i)%>" value="<%=arrData(1,i)%>" title="선택" />
										<label for=""><a href="javascript:openForm('<%=arrData(1,i)%>')" title="<%=arrData(2,i)%>"><%=FN_ClearTag(arrData(2,i))%></a></label></p>
								</li>
								<%			vNum = vNum - 1
										Next 
								%>
								<% End If %>
							</ul>
							</form>
						</div>
						
						<div class="paging">
							<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>
</html>