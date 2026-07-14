<!-- #include virtual="/common/CommonConfig.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : FileDown.asp / 파일 다운로드
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-20 / 강미현 / 4M / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%

	Dim Seq, Flag, Field, FName
	
	Call Main()
	
	Sub Main()
		Dim spName, spParam
		Dim strFileName

		Seq 	= FN_Req("dnSeq","0")
		Field 	= FN_Req("dnFld","")
		Flag 	= FN_Req("dnFlag","")
		FName 	= FN_Req("dnFName","")

		
		Select Case Flag
			Case "USERS"
				spName 	= "uspUser_Info"
				spParam = "@Seq"
			Case "BBS"
				spName 	= "uspBbs_Info"
				spParam = "@BbsSeq"
			Case "CONF"
				spName 	= "uspConfSite_Info"
				spParam = ""
			Case "STATIC"
			Case Else
				Call getError
				Response.End
		End Select
		
		
		If Flag = "STATIC" Then
			strFileName = "/_data/file_download/" & FN_Req("dnFName","")
		Else
			strFileName = getFileName(spName, spParam)
		End If

		
		
		If strFileName = "" Then
			' 사이트 설정의 첨부 파일이 없을경우.
			If Flag = "CONF" Then
				Call SB_ReturnErr("준비중입니다.","BACK")
				Response.end 
			Else
				Call getError
				Response.End
			End If 
		Else
			Call getFile(strFileName)
		End If
	
	End Sub
	
	Function getFileName(argSpName, argSpParam)
	
		Dim Rs
		Dim fnFileName : fnFileName = ""
		
		With objCmd
			.ActiveConnection = objDbCon
			.CommandType = 4
			.CommandText 					=argSpName
			.Parameters.Refresh
			.Parameters("@Flag").Value 		= "FDOWN"
			
			If argSpParam <> "" Then
			.Parameters(argSpParam).Value 	= Seq
			End If
			
			Set Rs = .Execute
		End With
			
		If Not Rs.Eof Then
			fnFileName = Rs(Field)
		End If
		
		Rs.Close
		Set Rs = Nothing
		
		getFileName = fnFileName
		
	End Function
	
	Sub getFile(argFilePath)

		Dim objFs, objF, objStream, buff, i
		Dim FilePath, FileName, FileEx
		Dim clsObjUpload : clsObjUpload = null
		
		
		FilePath = Replace(argFilePath,"..","")

	
		If instr(FilePath, "..") > 0 OR instr(FileName, "..") > 0 Then
			Call getError
			Response.End
		End If

		' 파일명 확인		
		If Trim(FName) = "" Then
			FileName = Mid(FilePath, InStrRev(FilePath, "/")+1)
		Else
			FileName = FName
		End If
		
		'업로드 금지된 확장자
		Set clsObjUpload = New ClsFileUpload
		If Not clsObjUpload.CheckExt(FileName) Then
			Call getError
			Response.end
		End If
		If Not IsNull(clsObjUpload) Then 	Set clsObjUpload = Nothing

		Set objFS = Server.CreateObject("Scripting.FileSystemObject")
		


		'파일이 존재 하는가?
		If NOT objFS.FileExists(Server.MapPath(FilePath)) Then
			Call SB_ReturnErr("요청하신 파일이 존재하지 않습니다.","BACK")
			Response.end 
		End If
		Set objFS = nothing
		
		Response.Clear
		Response.ContentType = "application/octet-stream" 
		Response.AddHeader "Content-Disposition","attachment; filename=" & server.URLPathEncode(FileName) 
		'Response.AddHeader "Content-Disposition","attachment; filename=" & FileName
		Response.AddHeader "Content-Transfer-Encoding","binary"
		Response.AddHeader "Expires","0"
		
		If Request.ServerVariables("HTTPS") = "off" Then
			'https 에서는 no-cashe 일경우 동적 파일다운로드 불가.
			Response.AddHeader "Pragma","no-cache"
		End If
		
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1
		objStream.LoadFromFile(Server.MapPath(FilePath))
	
		buff = objStream.Read

		Response.BinaryWrite buff 
	
		ObjStream.Close
		Set objStream = nothing
		Response.End	
	End Sub
	
	Sub getError()
		Call SB_ReturnErr("잘못된 경로로 접근하셨습니다.","BACK")
	End Sub
%>
