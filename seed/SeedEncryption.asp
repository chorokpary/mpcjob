<%@ codepage="65001" language="VBScript" %>
<!-- #include virtual="/common/Seed/config.asp" -->
<%
  ' 키파일(key.dat) 위치정보 include
	' 유니코드로 파라미터가 전송되므로 값을 코드페이지를 유니코드로 설정
	' 0 : ANSI (기본값)
	' 949 : 한국어 (EUC-KR)
	' 65001 : 유니코드 (UTF-8)
	' 65535 : 유니코드 (UTF-16)
	sPlainText = Trim( Request("iPlainText") )

	set oSeed = Server.CreateObject( "seed.CBC" )
	returnMsg = oSeed.LoadConfig(KEY_PATH)
	if StrComp("OK", returnMsg)= 0 then
		Response.Write( oSeed.Encrypt(sPlainText) )
	else
		Response.Write( returnMsg )
	end if
	
	set oSeed = nothing
%>
