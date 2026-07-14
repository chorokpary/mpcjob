<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_memo.asp / 관리자 메모 팝업
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
			.CommandText ="uspUserMemo_List"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "ADM"
			.Parameters("@pno").Value 		= Page
			.Parameters("@per_page").Value 	= listRecord
			.Parameters("@UsrSeq").Value 	= Seq
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			Set Rs = .Execute
		End With
			
		If Rs.Eof or Rs.Bof Then
			arrDataNum 	= -1
			kData 		= False
		Else
			' TotalCnt(0), UsrMSeq(1), AdmName(2), Memo(3), RegDate(4)
			kFields 	= Array("TotalCnt","UsrMSeq","AdmName","Memo","RegDate")
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
	<title>관리자 메모 - HC 관리자</title>
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
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); })
			$("#btnSubmit").click(chkForm);
		});
	//-->
	</script>
</head>

<body class="pop">
	<div id="wrap" class="Memo">

		<div id="contArea">
			<h2>관리자 메모 - <span class="weight_n"><%=StrUsrName%></span></h2>
			<div class="tblArea">
				<form name="frmPage" id="frmPage" method="post">
				<input type="hidden" id="Page" name="Page" value="<%=Page%>">
				</form>
				<table>
					<colgroup>
						<col class="w65" />
						<col />
						<col class="w75" />
						<col class="w50" />
					</colgroup>
					<thead>
						<tr class="design">
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<tr>
							<th>작성자</th>
							<th>내용</th>
							<th>작성일자</th>
							<th>삭제</th>
						</tr>
					</thead>
					<tbody>
						<% If Not kData Then %>
						<tr>
							<td colspan="4">등록된 메모가 없습니다.</td>
						</tr>
						<% Else %>
						<% 		
								' TotalCnt(0), UsrMSeq(1), AdmName(2), Memo(3), RegDate(4)
								vNum = RecordCount - ((Page-1) * listRecord) 
								For i = 0 To arrDataNum 
						%>
						<tr>
							<td><%=arrData(2,i)%></td>
							<td class="alignL"><p><%=FN_BrToNL(arrData(3,i))%></p></td>
							<td><%=FN_SetDateTimeFormat(arrData(4,i),"YYYY-MM-DD")%></td>
							<td><a href="javascript:doDelete(<%=arrData(1,i)%>)"><img src="/admin/images/btn_del.jpg" alt="삭제" /></a></td>
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
			
			<form id="frmInfo" name="frmInfo">
			<input type="hidden" id="Flag" name="Flag" value="MEMO">
			<input type="hidden" id="FlagSub" name="FlagSub">
			<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
			<input type="hidden" id="Sort" name="Sort">
			<input type="hidden" id="IsAjax" name="IsAjax" value="0">
			
			<h2>관리자 메모 추가</h2>
			<div class="inputArea">
				<textarea id="Memo" name="Memo"></textarea>
			</div>
			</form>
			<div class="btnArea">
				<div class="fl_l">
					<span class="btns btn_b"><a href="#" id="btnSubmit">등록</a>
					</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="#btnClose" id="btnClose">x 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>