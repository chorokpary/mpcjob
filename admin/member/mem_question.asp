<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : mem_memo.asp / 사전질문 답변 팝업
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2014-10-01 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim listRecord, listPage, Page, RecordCount, PageCount
	Dim sKey, sWord, Seq, RecrSeq
	Dim arrData, arrDataNum, kData, kFields, i, vNum
	
	Dim StrUsrName, StrPjtName

	Call Main()
	
	Sub Main()

		listPage 	= 10
		listRecord 	= 10	'목록갯수
		Page 		= FN_Req("Page","1")


		Seq 	= FN_Req("Seq","0")
		RecrSeq 	= FN_Req("RecrSeq","0")
		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If IsNumeric( Page ) Then
			Page = CInt( Page )
		Else
			Page = 1
		End If
		
		Call getList()
		StrUsrName = getName(Seq)
		StrPjtName = getRecruitInfo(RecrSeq)
		
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

	Function getRecruitInfo(argSeq)

		Dim Rs, StrTitle
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText ="uspRecruit_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "INFO"
			.Parameters("@Seq").Value 		= argSeq

			Set Rs = .Execute
		End With
		
		If Rs.Eof or Rs.Bof Then
			'권한 없음
			Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","CLOSE")
		Else
			StrTitle = "[" & Rs("PrjName") & "] " & Rs("Title")
		End If
	
		Rs.Close
		Set Rs = Nothing
		
		getRecruitInfo = StrTitle

	End Function	
		
	Sub getList()

		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText = "uspRecruitQaQ_Info"
			.Parameters.Refresh
			.Parameters("@Flag").Value 	= "LST_FRONT"
			.Parameters("@Seq").Value 	= RecrSeq
			.Parameters("@UsrSeq").Value = Seq
			
			' ## 회원정보
			Set Rs = .Execute
	
			If Not Rs.Eof Then
				kFields 	= Array("QSeq","QCode","Sort","Question","IsDisp","Answer")
				arrData = Rs.GetRows(,,kFields)
				arrDataNum = UBound(arrData, 2)
			Else
				arrData = null
			End If
		End With	
		
		Rs.Close
		Set Rs = Nothing

	End Sub
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>사전질문 답변 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">
	<!--
		function chkForm(){
			var len = $("#DtSize").val();
			var ckForm = true;
			
			if (len==0){
				alert("등록할 답변이 없습니다.");
			}else{
				/*
				// 유효성체크
				$("#frmInfo td.a textarea").each(function(){
					if (!$(this).val()){
						alert("질문의 답변을 입력해주세요.");
						$(this).focus();
						ckForm = false;
						return false;
					}
				});
				*/
			}
			
			if (ckForm) {
				$("#frmInfo").attr({action:"mem_proc.asp", method:"post"}).submit();
			}
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
			<h2>사전질문 답변</h2>
			<p class="note_c"><%=StrUsrName%> / <%=StrPjtName%></p>
			<div class="tblArea">
				<form id="frmInfo" name="frmInfo">
					<input type="hidden" id="Flag" name="Flag" value="QUES">
					<input type="hidden" id="FlagSub" name="FlagSub">
					<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
					<input type="hidden" id="RecrSeq" name="RecrSeq" value="<%=RecrSeq%>">
					<input type="hidden" id="DtSize" name="DtSize" value="<%=arrDataNum+1%>">
					<input type="hidden" id="IsAjax" name="IsAjax" value="0">
	
					<table>
						<colgroup>
							<col class="w80" />
							<col />
						</colgroup>
						<thead>
							<tr class="design">
								<td></td>
								<td></td>
							</tr>
						</thead>
						<tbody>
							<% 
							If Not IsArray(arrData) Then 
							%>
							<tr>
								<td colspan="2">등록된 질문이 없습니다.</td>
							</tr>
							<% 
							Else 
								' QSeq(0), QCode(1), Sort(2), Question(3), IsDisp(4), Answer(5)
								For i = 0 To arrDataNum 
									vNum = i + 1
							%>
							<tr>
								<th>질문 <%=vNum%></th>
								<td class="alignL"><%=arrData(3,i)%></td>
							</tr>
							<tr>
								<td colspan="2" class="a">
									<textarea id="Answer_<%=vNum%>" name="Answer_<%=vNum%>" style="width:870px;"><%=(arrData(5,i))%></textarea>
									<input type="hidden" name="QCode_<%=vNum%>" id="QCode_<%=vNum%>" value="<%=FN_InputVal(arrData(1,i))%>" >
								</td>
							</tr>
							<%
								Next 
							End If
							%>
						</tbody>
					</table>
				</div>
			</form>
			<div class="btnArea mb20">
				<div class="fl_l">
					<span class="btns btn_b"><a href="javascript:;" id="btnSubmit">등록</a>
					</span><span class="btns btn_g"><a href="javascript:;" id="btnCancel">취소</a></span>
				</div>
				<div class="fl_r">
					<span class="btns btn_g"><a href="javascript:;" id="btnClose">x 닫기</a></span>
				</div>
			</div>
		</div>
	</div>
</body>
</html>