<%@ LANGUAGE="VBScript" CODEPAGE=65001 LCID=1042 ENABLESESSIONSTATE=TRUE %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 
	'Option Explicit
	Response.CharSet = "UTF-8"
	
	Dim FUrl : FUrl = "/_data/SmartEditor/"
	Dim DefaultPath : DefaultPath = server.mappath(FUrl)
	FUrl = "http://www.mpcjob.co.kr/_data/SmartEditor/"

	
	Dim uploadform
	Set uploadform = Server.CreateObject("TABSUpload4.Upload")
	uploadform.CodePage = 65001
	uploadform.MaxBytesToAbort = 10485760
	uploadform.Start DefaultPath
	
    Dim callback_func : callback_func = uploadform("callback_func") '// 팝업창에 생성하는 iframe 이름 입니다. 이 값은 그대로 받아서 그대로 넘김니다. 
    
    Dim objImg : Set objImg = uploadform.Form("uploadInputBox")
    
	If objImg.FileSize > uploadform.MaxBytesToAbort Then
	   Page_Msg_Back "10M 이상의 사이즈인 파일은 업로드하실 수 없습니다"
	   Response.end
	End If
	
	'mimetype = uploadform.Form("uploadInputBox").ContentType
	'If mimetype <> "image/pjpeg" AND mimetype <> "image/gif" Then
	'   Page_Msg_Back "이미지만 업로드 할수 있습니다.[" & mimetype & "]"
	'   Response.end
	'End If
	
	imagepath = trim(uploadform("imagepath"))
	irid = trim(uploadform("id"))
	
	strFileName = uploadform("uploadInputBox").FileName
'	strFileWholePath = GetUniqueName(strFileName, server.mappath(FUrl))
'	uploadform("uploadInputBox").SaveAs strFileWholePath

	Call objImg.Save(DefaultPath ,False)
	strFileName = objImg.ShortSaveName

'	Set fs = server.CreateObject("Scripting.FileSystemObject")
'	strFileName = fs.GetFileName(strFileWholePath)
	'//////////////////////////////////////////////////////////////////////////////////
	'// 파일관련
	'////////////////////////////////// ///////////////////////////////////////////////
	
	 Function GetUniqueName(byRef strFileName, DirectoryPath)
	     Dim strName, strExt
	     strName = Mid(strFileName, 1, InstrRev(strFileName, ".") - 1)
	     strExt  = Mid(strFileName, InstrRev(strFileName, ".") + 1)
	
	     Dim fso
	     Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	     Dim bExist : bExist = True
	     '우선 같은이름의 파일이 존재한다고 가정
	     Dim strFileWholePath : strFileWholePath = DirectoryPath & "\" & strName & "." & strExt
	     '저장할 파일의 완전한 이름(완전한 물리적인 경로) 구성
	     Dim countFileName : countFileName = 0
	     '파일이 존재할 경우, 이름 뒤에 붙일 숫자를 세팅함.
	
	     Do While bExist ' 우선 있다고 생각함.
	         If (fso.FileExists(strFileWholePath)) Then ' 같은 이름의 파일이 있을 때
	             countFileName = countFileName + 1 '파일명에 숫자를 붙인 새로운 파일 이름 생성
	             strFileName = strName & "(" & countFileName & ")." & strExt
	             strFileWholePath = DirectoryPath & "\" & strFileName
	         Else
	             bExist = False
	         End If
	     Loop
	     GetUniqueName = strFileWholePath
	 End Function
	
	Sub Page_Msg_Back(msg)
	   with response
	      .write "<script language='javascript'>" & vbNewLine
	      .write "  alert('" & msg & "');" & vbNewLine
'	      .write "  history.back();" & vbNewLine
	      .write "</script>" & vbNewLine
	   End with
	End Sub
   
    '// 이부분이 중요합니다. 
    '// 파일업로드 처리하고 나서 callback.html로 업로드된 파일 정보를 넘겨줍니다. 
    '// callback_func 은 위에서 말씀드렸듯이 iframe이름인데 받아서 그대로 넘겨줍니다. 
    '// bNewLine=true 이 값은 제생각엔 true로 넘겨주면 에디터에서 이미지가 붙고 이미지 밑으로 커서를 내린다는 뜻인듯 합니다. 
    '// sFileName 은 이미지 파일명 
    '// sFileURL 은 이미지 URL 

    response.redirect "callback.html?callback_func="&callback_func&"&bNewLine=true&sFileName="&Server.URLEncode(strFileName)&"&sFileURL="&FUrl & Server.URLEncode(strFileName)
'    response.write "callback.html?callback_func="&callback_func&"&bNewLine=true&sFileName="&Server.URLEncode(strFileName)&"&sFileURL="&FUrl & Server.URLEncode(strFileName)
 %> 
