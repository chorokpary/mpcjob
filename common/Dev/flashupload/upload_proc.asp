<!-- #include virtual="/common/include/config.asp" -->
<%
'____________________________________________________________________________________
'
' * Discription : upload_proc.asp / 에디터 이미지 업로드
' 
' * History : (position number : date / author /회사약칭/ explanation)
'   #000 : 2010-11-01 / 강미현 / 4m / 최초작성 - 새폴더 생성 권한 없을경우 DB 연동방식으로 변경
'   #001 : 
'   #002 :
'   #003 :
'____________________________________________________________________________________
%>
<%
	Dim arrFiles
	
	arrFiles = FN_FileUpload("/smarteditor/" & Session("DirPath"),"Filedata",True, True)

%>