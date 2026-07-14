<!-- #include virtual="/common/include/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : upload_filelist.asp / 에디터 이미지 업로드 - 파일 목록
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2010-11-09 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Response.ContentType = "text/xml"
	
	Dim fs, folderObj, files, item

	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	
	Set folderObj = fs.GetFolder(Server.MapPath("/fileupload/smarteditor/" & Session("DirPath")))
	Set files = folderObj.Files 
%>
	<files>
<%		
	For Each item in files
		%>
		<url>/<%=Session("DirPath")%>/<%=item.name%></url>	
		<%
   Next
%>
	</files>
