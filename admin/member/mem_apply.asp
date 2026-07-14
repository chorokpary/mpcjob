<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_apply.asp / 개별회원 입사지원 현황 팝업
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-21 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, Seq
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim StrUsrName
	Dim strAppCStyle, strAppCMsg

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 10	'목록갯수
		Page 		= FN_Req("Page","1")


		Seq 	= FN_Req("Seq","0")
		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		StrUsrName = getName(Seq)
		
	End Sub

	Function getName(argSeq)

		Dim Rs, Rtn
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspUser_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "SIMP"
			.Parameters("@Seq").Value 	= Seq
			
			Set Rs = .Execute
		End With
			
		If Not Rs.Eof Then
			Rtn = Rs("UsrName")
		Else
			Rtn = ""
		End If
		
		Rs.Close
		Set Rs = Nothing
		
		getName = Rtn

	End Function
		
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspApply_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM_U"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@Seq").Value 	= Seq
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), AppSeq(1), RecrSeq(2), RegDate(3), Status(4), PrjName(5), Title(6), AppBDate(7), AppEDate(8), PartnerDepth1(9), PartnerDepth2(10), CancelDate(11)
			kFields 	= Array("TotalCnt","AppSeq","RecrSeq","RegDate","Status","PrjName","Title","AppBDate","AppEDate","PartnerDepth1","PartnerDepth2","CancelDate")
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
	<title>입사지원내역 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			if (!$("#Memo").val()){
				alert("메모를 입력해주세요.");
				$("#Memo").focus();
				return;
			}
			if ($("#Memo").val().length > 60){
				alert("메모는 최대 60자까지 입력 할 수 있습니다.");
				$("#Memo").focus();
				return;
			}
			$("#FlagSub").val("ADD");
			$("#frmInfo").attr({action:"mem_proc.asp", method:"post"}).submit();
		}
		
		function doDelete(seq){
			$("#FlagSub").val("DEL");
			$("#Sort").val(seq);
			$("#frmInfo").attr({action:"mem_proc.asp", method:"post"}).submit();
		}
		
		$(function() {
		
			$("#btnUserInfo").click(function(){ openWin("../member/mem_print.asp?Seq=<%=Seq%>","popPrint",0,0,930,650,1); });
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); })
			$("#btnSubmit").click(chkForm);
		});
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap" class="applicant">
		<div id="contArea">
			<h2><%=StrUsrName%> 님의 입사지원내역</h2>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#btnUserInfo" id="btnUserInfo">회원 이력서 바로가기</a></span>
				</div>
			</div>
			<div class="tblArea">
				<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				</form>
				
				<table>
					<colgroup>
						<col class="w50" />
						<col class="w120" />
						<col />
						<col class="w100" />
						<col class="w100" />
						<col class="w165" />
					</colgroup>
					<thead>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<tr>
							<th>번호</th>
							<th>프로젝트명</th>
							<th>공고제목</th>
							<th>지원경로</th>
							<th>상태</th>
							<th>채용기간</th>
						</tr>
					</thead>
					<tbody>
						<% If Not kData Then %>
						<tr>
							<td colspan="6">지원 내역이 없습니다.</td>
						</tr>
						<% Else %>
						<% 		
								' TotalCnt(0), AppSeq(1), RecrSeq(2), RegDate(3), Status(4), PrjName(5), Title(6), AppBDate(7), AppEDate(8), PartnerDepth2(10), CancelDate(11)
								vNum = RecordCount - ((Page-1) * listRecord) 
								For i = 0 To arrDataNum 
									strAppCStyle = "" : strAppCMsg = ""

									If Not IsNull(arrData(11,i)) Then
										strAppCStyle = " class=""cancel"""
										strAppCMsg = "<br>/ 지원취소"
									End If
						%>
						<tr<%=strAppCStyle%>>
							<td><%=vNum%></td>
							<td><%=arrData(5,i)%></td>
							<td class="alignL"><%=arrData(6,i)%></td>
							<td><%=arrData(9,i)%> 
								<% If Not FN_IsBlank(arrData(10,i)) Then %><br>(<%=arrData(10,i)%>)<% End If %>
							</td>
							<td><%=arrData(4,i)%><%=strAppCMsg%></td>
							<td><%=FN_SetDateTimeFormat(arrData(7,i),"YYYY-MM-DD")%> ~ <%=FN_SetDateTimeFormat(arrData(8,i),"YYYY-MM-DD")%></td>
						</tr>
						<%			vNum = vNum - 1
								Next 
						%>
						<% End If %>
					</tbody>
				</table>
				<div class="paging">
					<% Call SB_PagingWriteAdmin(listPage, RecordCount, PageCount) %>
				</div>
			</div>
			<div class="btnArea">
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">X 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>