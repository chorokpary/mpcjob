<%
	Dim tmpDt
	tmpDt = NOW()
	Session("DirPath") = Year(tmpDt) & "/" &_
			Year(tmpDt) &_
			Right("00" & Month(tmpDt),2) &_
			Right("00" & Day(tmpDt),2) &_
			Right("00" & Hour(tmpDt),2) &_
			Right("00" & Minute(tmpDt),2) &_
			Right("00" & Second(tmpDt),2) &_
			"_" & Session.SessionID
			
			response.write session("DirPath")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<html>
	<head>
		<meta http-equiv="Content-type" content="text/html; charset=utf-8">
		<meta http-equiv="cache-control" content="no-cache, must-revalidate">
		<meta http-equiv="pragma" content="no-cache">
		<title>이미지삽입</title>
		<link rel="stylesheet" type="text/css" href="/common/SmartEditor/css/editor.css" />
		<link href="upload.css" rel="stylesheet" type="text/css" />

	<link href="browser.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" src="js/common.js"></script>
	<script type="text/javascript">

var oListManager = new Object() ;

oListManager.Clear = function()
{
	document.body.innerHTML = '' ;
}

oListManager.GetFolderRowHtml = function( folderName, folderPath )
{
	// Build the link to view the folder.
	var sLink = '<a href="#" onclick="OpenFolder(\'' + folderPath.replace( /'/g, '\\\'') + '\');return false;">' ;

	return '<tr>' +
			'<td width="16">' +
				sLink +
				'<img alt="" src="images/Folder.gif" width="16" height="16" border="0"></a>' +
			'</td><td nowrap colspan="2">&nbsp;' +
				sLink +
				folderName +
				'</a>' +
		'</td></tr>' ;
}

oListManager.GetFileRowHtml = function( fileName, fileUrl, fileSize )
{
	// Build the link to view the folder.
	var sLink = '<a href="#" onclick="OpenFile(\'' + fileUrl.replace( /'/g, '\\\'') + '\');return false;">' ;

	// Get the file icon.
	var sIcon = oIcons.GetIcon( fileName ) ;

	return '<tr>' +
			'<td width="16">' +
				sLink +
				'<img alt="" src="images/icons/' + sIcon + '.gif" width="16" height="16" border="0"></a>' +
			'</td><td>&nbsp;' +
				sLink +
				fileName +
				'</a>' +
			'</td><td align="right" nowrap>&nbsp;' +
				fileSize +
				' KB' +
		'</td></tr>' ;
}

function OpenFolder( folderPath )
{
	// Load the resources list for this folder.
	window.parent.frames['frmFolders'].LoadFolders( folderPath ) ;
}

function OpenFile( fileUrl )
{
	window.top.opener.SetUrl( encodeURI( fileUrl ) ) ;
	window.top.close() ;
	window.top.opener.focus() ;
}

function LoadResources( resourceType, folderPath )
{
	oListManager.Clear() ;
	oConnector.ResourceType = resourceType ;
	oConnector.CurrentFolder = folderPath ;
	oConnector.SendCommand( 'GetFoldersAndFiles', null, GetFoldersAndFilesCallBack ) ;
}

function Refresh()
{
	LoadResources( oConnector.ResourceType, oConnector.CurrentFolder ) ;
}

function GetFoldersAndFilesCallBack( fckXml )
{
	if ( oConnector.CheckError( fckXml ) != 0 )
		return ;

	// Get the current folder path.
	var oFolderNode = fckXml.SelectSingleNode( 'Connector/CurrentFolder' ) ;
	if ( oFolderNode == null )
	{
		alert( 'The server didn\'t reply with a proper XML data. Please check your configuration.' ) ;
		return ;
	}
	var sCurrentFolderPath	= oFolderNode.attributes.getNamedItem('path').value ;
	var sCurrentFolderUrl	= oFolderNode.attributes.getNamedItem('url').value ;

//	var dTimer = new Date() ;

	var oHtml = new StringBuilder( '<table id="tableFiles" cellspacing="1" cellpadding="0" width="100%" border="0">' ) ;

	// Add the Folders.
	var oNodes ;
	oNodes = fckXml.SelectNodes( 'Connector/Folders/Folder' ) ;
	for ( var i = 0 ; i < oNodes.length ; i++ )
	{
		var sFolderName = oNodes[i].attributes.getNamedItem('name').value ;
		oHtml.Append( oListManager.GetFolderRowHtml( sFolderName, sCurrentFolderPath + sFolderName + "/" ) ) ;
	}

	// Add the Files.
	oNodes = fckXml.SelectNodes( 'Connector/Files/File' ) ;
	for ( var j = 0 ; j < oNodes.length ; j++ )
	{
		var oNode = oNodes[j] ;
		var sFileName = oNode.attributes.getNamedItem('name').value ;
		var sFileSize = oNode.attributes.getNamedItem('size').value ;

		// Get the optional "url" attribute. If not available, build the url.
		var oFileUrlAtt = oNodes[j].attributes.getNamedItem('url') ;
		var sFileUrl = oFileUrlAtt != null ? oFileUrlAtt.value : sCurrentFolderUrl + sFileName ;

		oHtml.Append( oListManager.GetFileRowHtml( sFileName, sFileUrl, sFileSize ) ) ;
	}

	oHtml.Append( '<\/table>' ) ;

	document.body.innerHTML = oHtml.ToString() ;

//	window.top.document.title = 'Finished processing in ' + ( ( ( new Date() ) - dTimer ) / 1000 ) + ' seconds' ;

}

window.onload = function()
{
	window.top.IsLoadedResourcesList = true ;
}
	</script>

	</head>
	<body>
		<div id="naver_common_editor">
			<form id="editor_upimage" name="editor_upimage" action="upload_proc.asp" method="post" enctype="multipart/form-data">
			<input type="hidden" name="id" value="<%=request("id")%>">
			<input type="file" name="Filedata" id="Filedata">
			<input type="submit">
			</form>
		</div>
	</body>
</html>
