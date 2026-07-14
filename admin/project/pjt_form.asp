<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : pjt_form.asp / 프로젝트 등록, 수정
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-06-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag
	Dim PrjName, LogoPath, MapPath, Addr, Part, PrjFlag
	Dim TransSubway, TransBus, Location, Center, IsView
	
	Call Main()
	
	Sub Main()

		Seq = FN_Req("Seq","0")
		
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
			.CommandText ="uspProject_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "INFO"
			.Parameters("@Seq").Value 	= Seq
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PrjName = FN_InputVal(Rs("PrjName"))
				LogoPath = Rs("LogoPath")
				MapPath = Rs("MapPath")
				Addr 	= FN_InputVal(Rs("Addr"))
				Part 	= Rs("Part")
				PrjFlag = Rs("Flag")
				TransSubway = FN_InputVal(Rs("TransSubway"))
				TransBus 	= Rs("TransBus")
				Location 	= Rs("Location")
				Center 		= Rs("Center")
			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>프로젝트 등록 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			if (!$("#PrjName").val()){
				alert("프로젝트명을 입력해주세요.");
				$("#PrjName").focus();
				return;
			}
		
			$("#frmInfo").attr({action:"pjt_proc.asp", method:"post"}).submit();
		}
		
		function doDelFile(targetImg, targetFld, strPath){
			$("#" + targetImg).hide();
			$("#" + targetFld).val(strPath);
		}
		
		$(function() {
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); $("#pLogoPath").show(); $("#pMapPath").show(); })
			$("#btnSubmit").click(chkForm);
			
			
			$("#TransBus").keyup(function(){ limitCharacters(this.id, 1000 ,'') });
		});
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap">
		<div id="contArea">
			<h2>프로젝트 등록</h2>
			<div class="tblArea">

				<form id="frmInfo" name="frmInfo" action="pjt_proc.asp" method="post" enctype="multipart/form-data" >
				<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
				<input type="hidden" id="Flag" name="Flag" value="<%=Flag%>">
				<table class="tblForm">
					<colgroup>
						<col class="w85" />
						<col class="w85" />
						<col class="w125" />
						<col class="w85" />
						<col />
						<col class="w85" />
						<col />
					</colgroup>
					<tbody>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<tr>
							<th><label for="PrjName">프로젝트명</label></th>
							<td colspan="6">
								<input type="text" id="PrjName" name="PrjName" Title="프로젝트명" value="<%=PrjName%>" maxlength="15" class="w780" /> 
							</td>
						</tr>
						<tr>
							<th><label for="Addr">주소</label></th>
							<td colspan="6">
								<input type="text" id="Addr" name="Addr" Title="주소" value="<%=Addr%>" maxlength="150" class="w780" /> 
							</td>
						</tr>
						<tr>
							<th><label for="Part">분야</label></th>
							<td colspan="2">
								<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_PROJECT", "Part", "Part", "", "", Part, "등록된 분야가 없습니다.", " class=""w160"" ", False)%>
							</td>
							<th><label for="PrjFlag">모집구분</label></th>
							<td>
								<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_PROJECT", "Flag", "PrjFlag", "", "", PrjFlag, "등록된 분야가 없습니다.", " class=""w160"" ", False)%>
							</td>
							<th><label for="Center">센터선택</label></th>
							<td>
								<% Call SB_MakeHTMLSelect_CdTbl(objDBCon, "TBL_PROJECT", "Center", "Center", "", "", Center, "등록된 센터가 없습니다.", " class=""w160"" ", False)%>
							</td>

						</tr>
						<tr>
							<th><label for="LogoPath">썸네일</label></th>
							<td colspan="6">
								<% If Not FN_isBlank(LogoPath) Then %>
								<p id="pLogoPath"><img src="<%=LogoPath%>" alt="" height="35" /><span class="btns btn_in"><a href="javascript:doDelFile('pLogoPath','DelLogoPath','<%=LogoPath%>')">삭제</a></span></p>
								<% End If %>
								<input type="file" id="LogoPath" name="LogoPath" title="썸네일" value="" class="w675" size="95" /> (최적화 : 75 * 75)
								<input type="hidden" id="DelLogoPath" name="DelLogoPath" value="">
							</td>
						</tr>
						<tr>
							<th rowspan="2"><label for="TransSubway">교통편</label></th>
							<td class="traffic">지하철</td>
							<td colspan="5"><input type="text" id="TransSubway" name="TransSubway" title="지하철" value="<%=TransSubway%>" maxlength="100" class="w695" /></td>
						</tr>
						<tr>
							<td class="traffic">버스</td>
							<td colspan="5"><textarea id="TransBus" name="TransBus" title="버스" class="pjtTxt" maxlength="500"><%=TransBus%></textarea></td>
						</tr>
						<tr>
							<th><label for="Location">오시는길</label></th>
							<td colspan="6">
								<textarea id="Location" name="Location" title="오시는길"><%=Location%></textarea>
							</td>
						</tr>
						<tr>
							<th><label for="MapPath">약도등록</label></th>
							<td colspan="6">
								<% If Not FN_isBlank(MapPath) Then %>
								<p id="pMapPath"><%=Mid(MapPath,InStrRev(MapPath,"/")+1)%> <span class="btns btn_in"><a href="javascript:doDelFile('pMapPath','DelMapPath','<%=MapPath%>')">삭제</a></span></p>
								<% End If %>
								<input type="file" id="MapPath" name="MapPath" title="약도등록" value="" class="w675" size="95" /> (최적화 : 726 * 262)
								<input type="hidden" id="DelMapPath" name="DelMapPath" value="">
							</td>
						</tr>
					</tbody>
				</table>
				</form>
			</div>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#" id="btnSubmit"><%=Fn_SetDefault(Flag,"ADD","등록하기","수정하기")%></a>
					</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소하기</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">x 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>