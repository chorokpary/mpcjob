<%
'____________________________________________________________________________________
'
' * Discription : Class.FileUpload.asp / 파일 업로드 컴포넌트 관리
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2012-07-05 / 강미현 / 4m / 최초작성
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________


Class ClsFileUpload

	Private StrErrorMsg
	Private ObjUpload
	Private Const_ObjType, Const_MaxSize, Const_Path
	Private StrFileName

	'========================================================
	' 클래스 생성자.
	'========================================================
    Sub Class_Initialize()
		StrErrorMsg = ""
		ObjUpload = null
		Const_ObjType = ""
		Const_Path = ""
		Const_MaxSize = 0
    End Sub


	'========================================================
	' 클래스 소멸자
	'========================================================
    Sub Class_Terminate
    	Set ObjUpload = Nothing
    End Sub
    

	'========================================================
	' SET CreateObject
	'========================================================	
	Public Function setCreateObj (argObjFlag, argDefPath, argMaxSize)
	
		Select Case argObjFlag
			Case "DEXT"
				'Set ObjUpload = Server.CreateObject("DEXT.FileUpload")
				
				ObjUpload.AutoMakeFolder = True
				ObjUpload.DefaultPath = argDefPath
				ObjUpload.MaxFileLen = argMaxSize
			Case "ABC" 
				Set ObjUpload = Server.CreateObject("ABCUpload4.XForm")
				
				ObjUpload.AbsolutePath = True
				ObjUpload.Overwrite = True
			Case "TABS"
				Set ObjUpload = Server.CreateObject("TABSUpload4.Upload")
				
				ObjUpload.CodePage = 65001
				ObjUpload.MaxBytesToAbort = argMaxSize
				ObjUpload.Start argDefPath
			Case "SG"
				Server.ScriptTimeout = 1000
				Set ObjUpload = Server.CreateObject("SiteGalaxyUpload.Form")
			Case Else
				setCreateObj = False
		End Select
		
		Const_MaxSize 	= argMaxSize
		Const_ObjType 	= argObjFlag
		Const_Path 		= argDefPath
		
		setCreateObj = True
		
	End Function

	
	'========================================================
	' GET Request Value 
	'========================================================	
	Public Function getReq(argName)
		Dim strVal : strVal = ObjUpload(argName)
		
		getReq = strVal
	End Function

	'========================================================
	' GET Save File State Clear
	'========================================================	
	Public Sub setClear()
	
	End Sub


	'========================================================
	' GET Save File
	'========================================================	
	Public Function SaveFile(argPath, argName)
	
		Dim ChkExt : ChkExt = CheckExt(ObjUpload(argName))
		
		' 확장자 체크
		If Not ChkExt Then 
			setErrMsg("업로드 할 수 없는 확장자 입니다.")
			SaveFile = False
			Response.End
		End If
		
		On Error Resume next

		' 파일 업로드
		Select Case Const_ObjType
			Case "DEXT" : 	StrFileName = SaveFile_Dext(argPath, argName)
			Case "ABC" : 	StrFileName = SaveFile_ABC(argPath, argName)
			Case "TABS" : 	StrFileName = SaveFile_TABS(argPath, argName)
			Case "SG" : 	StrFileName = SaveFile_SiteGalaxy(argPath, argName)
			Case Else :
				setErrMsg("업로드 컴포넌트 설정 환경을 확인해 주세요.")
				SaveFile = False
		End Select
		
		If Err.Numbe <> 0 Then
			'Response.Write Description
			setErrMsg("파일 업로드중 오류가 발생하였습니다.")
			SaveFile = False
			Err.Clear
		End If
		
		SaveFile = True
	End Function

	' +++ Save File : DextUpload
	Public Function SaveFile_Dext(argPath, argName)
		Call  ObjUpload(argName).Save(Const_Path + "\" + argPath ,False)
		SaveFile_Dext = ObjUpload(argName).LastSavedFileName
	End Function

	' +++ Save File : ABCUpload
	Public Function SaveFile_ABC(argPath, argName)
		SaveFile_ABC = ""
	End Function

	' +++ Save File : TABSUpload
	Public Function SaveFile_TABS(argPath, argName)
		Call ObjUpload.Form(argName).Save(Const_Path + "\" + argPath ,False)
		SaveFile_TABS = ObjUpload.Form(argName).ShortSaveName
	End Function

	' +++ Save File : SiteGalaxyUpload
	Public Function SaveFile_SiteGalaxy(argPath, argName)
		SaveFile_SiteGalaxy = ""
	End Function

	'========================================================
	' Set Get FileName
	'========================================================	
	Public Function getFileName()
		getFileName = StrFileName
	End Function


	'========================================================
	' File Filtering
	'========================================================	
	Public Function CheckExt(argFName)
		Dim NotExt : NotExt = "asp,dll,exe,com,bat,aspx,php,html,htm,jsp,php,cgi,perl,pl"
		Dim ArrNotExt : ArrNotExt = Split(NotExt,",")
		Dim ArrFileExt : ArrFileExt = Split(argFName,".")
		Dim Rtn : Rtn = True
		Dim I, J
		
		For I = 1 To UBound(ArrFileExt)
			For J = 0 To UBound(ArrNotExt)
				If UCASE(ArrFileExt(I)) = UCASE(ArrNotExt(J)) Then
					Rtn = False
				End If
			Next
		Next
		
		CheckExt = Rtn
	End Function


	'========================================================
	' File Delete
	'========================================================	
	Public Sub Delete(argDelPath)
		Dim fnObjFSO : Set fnObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		If fnObjFSO.FileExists(argDelPath) Then 		 fnObjFSO.DeleteFile(argDelPath)		

		Set fnObjFSO = Nothing
	End Sub


	'========================================================
	' 에러정보 입출력
	'========================================================	
	Public Sub setErrMsg(argMsg)
		StrErrorMsg = argMsg
	End Sub

	Public Function getErrMsg()
		getErrMsg = StrErrorMsg
	End Function
    
End Class
%>