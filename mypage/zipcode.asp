<!-- #include virtual="/common/CommonConfig.asp" -->
<!-- #include virtual="/common/FrontConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : zipcode.asp / 우편번호검색 (ajax)
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-27 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim arrData, arrDataNum, kFields, i
	
	Dim tgSi, tgGu, tgZip1, tgZip2, tgAddr, tgAddr2
	Dim sKey
	
	Call Main()
	
	Sub Main()

		tgSi = FN_Req("tgSi","AddrSido")
		tgGu = FN_Req("tgGu","AddrGugun")
		tgZip1 = FN_Req("tgZip1","Zipcode1")
		tgZip2 = FN_Req("tgZip2","Zipcode2")
		tgAddr = FN_Req("tgAddr","Addr")
		tgAddr2 = FN_Req("tgAddr2","AddrDetail")

		sKey = FN_Req("sKey","")
		
		arrDataNum = -1
		
		If sKey <> "" Then
			Call getList
		End If
	
	End Sub
	
	Sub getList()
	
		Dim Rs
		
		With objCmd
			.ActiveConnection = objDBCon
			.CommandType = 4
			.CommandText ="uspZipcode_Info"
			.Parameters.Refresh
			.Parameters("@FLAG").Value 	= "SADDR"
			.Parameters("@sKey").Value 	= sKey
			
			' ## 회원정보
			Set Rs = .Execute

			If Not Rs.Eof Then
				arrDataNum = 0
				' ZIPCODE(0), SIDO(1), GUGUN(2), DONG(3), BUNJI(4)
				
				kFields 	= Array("ZIPCODE","SIDO","GUGUN","DONG","BUNJI")
				arrData 	= Rs.GetRows(,,kFields)
				arrDataNum 	= UBound(arrData, 2)

			End If

			Rs.Close
			Set Rs = Nothing
		End With
		
	End Sub
	
%>
	<script type="text/javascript">
	<!--
		function doFillAddr(argZip, argSido, argGugun, argDong, argBunji){
			$("#frmInfo").find("#<%=tgZip1%>").val(argZip.substring(0,3));
			$("#frmInfo").find("#<%=tgZip2%>").val(argZip.substring(4,7));
			$("#frmInfo").find("#<%=tgSi%>").val(argSido);
			$("#frmInfo").find("#<%=tgGu%>").val(argGugun);
			$("#frmInfo").find("#<%=tgAddr%>").val(argSido + " " + argGugun + " " + argDong);
			$("#frmInfo").find("#<%=tgAddr2%>").val(argBunji);
			
			closeLayer();
		}
		
		function doSearchAddr(){
			var param = $("#frmZip").serialize();
			openLayer("zipcode.asp",param, "w528");
		
			return false;
		}
	//-->
	</script>

	<div class="inner"><div class="inner2">
		<div class="popCont zipcode">
			<h2><img src="/images/pages/tit_zipcode.jpg" alt="우편번호 검색" /></h2>
			<p class="note">
				<img src="/images/pages/txt_zipCode.jpg" alt="찾고싶으신 주소의 동(읍/면) 지역을 입력해 주세요." />
			</p>
			
			<form id="frmZip" name="frmZip" method="post" action="zipcode.asp" onsubmit="return doSearchAddr();">
			<input type="hidden" id="tgSi" name="tgSi" value="<%=tgSi%>">
			<input type="hidden" id="tgGu" name="tgGu" value="<%=tgGu%>">
			<input type="hidden" id="tgZip1" name="tgZip1" value="<%=tgZip1%>">
			<input type="hidden" id="tgZip2" name="tgZip2" value="<%=tgZip2%>">
			<input type="hidden" id="tgAddr" name="tgAddr" value="<%=tgAddr%>">
			<input type="hidden" id="tgAddr2" name="tgAddr2" value="<%=tgAddr2%>">

			<div class="zipForm">
				<input type="text" name="sKey" id="sKey" value="<%=sKey%>" title="검색어 입력" class="w355" style="ime-mode:active;"/> 
				<input type="image" src="/images/pages/btn_search.jpg" alt="검색하기" />
				<p class="ex">예) 역삼동, 신당동, 종로1가, 상암동</p>
			</div>
			</form>
			
			<% If arrDataNum > -1 Then %>
			<p class="total">총 <strong><%=arrDataNum+1%></strong> 개의 우편번호가 검색 되었습니다.</p>
			<div class="addrChioce">
				<ul>
					<% 
					' ZIPCODE(0), SIDO(1), GUGUN(2), DONG(3), BUNJI(4)
					
					For i = 0 To arrDataNum 
					%>
					<li><a href="javascript:doFillAddr('<%=arrData(0,i)%>','<%=arrData(1,i)%>','<%=arrData(2,i)%>','<%=arrData(3,i)%>','<%=arrData(4,i)%>')"><%=arrData(0,i)%>&nbsp;<%=arrData(1,i)%>&nbsp;<%=arrData(2,i)%>&nbsp;<%=arrData(3,i)%>&nbsp;<%=arrData(4,i)%></a></li>
					<% Next %>
				</ul>
			</div>
			<p class="description"><span>주소를 선택하면 자동으로 입력 됩니다.</span></p>
		</div>
		<% End If %>
	</div></div>
