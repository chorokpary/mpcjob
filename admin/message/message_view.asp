<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/AdminConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : message_view.asp / 쪽지 상세보기
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-28 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim Seq, Flag, Page
	Dim sKey, sWord
	Dim PrjName, RecrSeq, SendName, SendSeq, Tel, RegDate, Memo
	Dim nextSeq, nextSubject, nextDate, prevSeq, prevSubject, prevDate
	
	Call Main()
	
	Sub Main()

		Seq 	= FN_Req("Seq","0")
		Page 	= FN_Req("Page","1")
		sKey 	= FN_Req("sKey","")
		sWord 	= FN_InputVal(FN_Req("sWord",""))
		
		If Seq = "0" Then
			Call SB_ReturnErr("[PARAM] 잘못된 경로로 접근하셨습니다.","BACK")
		Else
			Call getData
		End If
	
	End Sub
	
	Sub getData()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspNote_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 		= "INFO"
			.Parameters("@Seq").Value 		= Seq
			.Parameters("@RecvSeq").Value 	= Session("ASeq")
			
			.Parameters("@sKey").Value 		= sKey
			.Parameters("@sWord").Value 	= sWord
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				PrjName 	= Rs("PrjName")
				RecrSeq 	= Rs("RecrSeq")
				SendName 	= Rs("SendName")
				SendSeq 	= Rs("SendSeq")
				Tel 		= Rs("Tel")
				RegDate 	= Rs("RegDate")
				Memo 		= HTMLEncode(Rs("Memo"))
			End If

			Set Rs = Rs.NextRecordSet
			
			If Not Rs.Eof Then
				prevSeq = Rs("Seq")
				prevSubject = HTMLEncode(Rs("Memo"))
			End If
	
			Set Rs = Rs.NextRecordSet
			
			If Not Rs.Eof Then
				nextSeq = Rs("Seq")
				nextSubject = HTMLEncode(Rs("Memo"))
			End If
		
			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<title>쪽지 - HC 관리자</title>
	<!--#include virtual="/admin/include/head.asp"-->
	<script type="text/javascript">

		function chkForm(){
			if (!$("#Tel").val()){
				alert("핸드폰가 등록되지 않는글 입니다.");
				return;
			}
			if (!IsMobile($("#Tel").val(),"") && !IsMobile($("#Tel").val(),"-")){
				alert("올바른 핸드폰번호가 아닙니다.");
				return;
			}
			if (!$("#Answer").val()){
				alert("답변내용을 입력해주세요.");
				$("#Answer").focus();
				return;
			}

			$("#Flag").val("ANS");
			$("#frmInfo").attr({action:"message_proc_n.asp", method:"post"}).submit();
		}
		
		function doDel(){
			if(confirm("쪽지를 삭제하시겠습니까?")){
				$("#Flag").val("DEL_RECV");
				$("#frmInfo").attr({action:"message_proc.asp", method:"post"}).submit();
			}
		}
		
		function goView(seq){
			goViewCommon(seq, "message_view.asp");
		}
		
		function goList(){
			goListCommon("message_list.asp");
		}
		
		function doCheckByte(){
			limitCharactersByte($("#Answer"),2000,$("#divLimitTxt"),"L");
		}

		$(document).ready(function() {
			// ### Button Control
			$("#btnCancel").click(function(){ $("#frmInfo")[0].reset(); })
			$("#btnList").click(goList);
			$("#btnDel").click(doDel);
			$("#btnAnswer").click(chkForm);
			
			$("#Answer").keyup(doCheckByte);
			//$("#Answer").blur(doCheckByte);
		});
		
		$(window).load(doCheckByte);
	</script>
</head>

<body>
	<div id="wrap" class="message">
		<!--#include virtual="/admin/include/top.asp"-->
		
		<div class="contents">
			<!--#include virtual="/admin/include/left.asp"-->
			<div class="right">
				<div class="location">
					<a href="" class="home">HOME</a><span>쪽지함</span>
				</div>
				<div id="contArea">
					<h2>쪽지함</h2>
					<div class="tblArea">
						<form id="frmInfo" name="frmInfo" action="message_action.asp" method="post">
						<input type="hidden" id="Flag" name="Flag">
						<input type="hidden" id="Seq" name="Seq" value="<%=Seq%>">
						<input type="hidden" id="Tel" name="Tel" value="<%=Tel%>">
						<input type="hidden" id="RecvSeq" name="RecvSeq" value="<%=SendSeq%>">
						<input type="hidden" id="txtID" name="txtID" value="<%=Replace(Tel,"-","")%>">
						<table>
							<colgroup>
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
								<col class="w100" />
								<col />
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
									<th>프로젝트 명</th>
									<td colspan="5"><%=PrjName%></td>
								</tr>
								<tr>
									<th>작성자</th>
									<td><%=SendName%></td>
									<th>연락처</th>
									<td><%=Tel%></td>
									<th>작성일자</th>
									<td><%=FN_SetDateTimeFormat(RegDate,"YYYY.MM.DD")%></td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<th>내용</th>
									<td class="alignL" colspan="5"><p><%=Memo%></p></td>
								</tr>
								<tr>
									<th>답변하기</th>
									<td class="alignL" colspan="5"><textarea id="Answer" name="Answer"></textarea><div id="divLimitTxt"></div></td>
								</tr>
							</tbody>
						</table>
					</div>
					</form>
					<div class="btnArea">
						<div class="fl_l">
							<span class="btns btn_b"><a href="#btnList" id="btnList">목록보기</a>
							</span><span class="btns btn_b"><a href="/recruit/recruit_view.asp?Seq=<%=RecrSeq%>" target="_blank">공고</a>
							</span><span class="btns btn_b"><a href="/admin/recruit/job_skin_mod.asp?Flag=MOD&Seq=<%=RecrSeq%>">공고수정</a></span>
						</div>
						<div class="fl_r">
							<span class="btns btn_b"><a href="#btnAnswer" id="btnAnswer">답변발송</a>
							</span><span class="btns btn_g"><a href="#btnDel" id="btnDel">삭제</a>
							</span><span class="btns btn_g"><a href="#btnCancel" id="btnCancel">취소</a></span>
						</div>
					</div>
					
					<form name="frmPage" id="frmPage" method="post">
					<input type="hidden" id="Seq" name="Seq">
					<input type="hidden" id="Page" name="Page" value="<%=Page%>">
					<input type="hidden" id="sKey" name="sKey" value="<%=sKey%>">
					<input type="hidden" id="sWord" name="sWord" value="<%=sWord%>">
					</form>

					<div class="tblArea">
						<table>
							<colgroup>
								<col class="w100" />
								<col />
							</colgroup>
							<thead>
								<tr class="design">
									<td></td>
									<td></td>
								</tr>
							</thead>
							<tbody>
								<tr>
									<th>이전글</th>
									<td class="alignL">
										<% If prevSeq > 0 Then %>
										<a href="javascript:goView(<%=prevSeq%>);"><%=prevSubject%></a>
										<% Else %>
										이전글이 없습니다.
										<% End If %>
									</td>
								</tr>
								<tr>
									<th>다음글</th>
									<td class="alignL">
										<% If nextSeq > 0 Then %>
										<a href="javascript:goView(<%=nextSeq%>);"><%=nextSubject%></a>
										<% Else %>
										다음글이 없습니다.
										<% End If %>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>
		<!--#include virtual="/admin/include/footer.asp"-->
	</div>
</body>

</html>