
function flash(w,h,u,id) {
 var flash_tag = "";
 flash_tag = '<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0" width="' +w+ '" height="'+h+'" id="'+id+'">';
 flash_tag +='<param name="allowScriptAccess" value="sameDomain" />';
 flash_tag +='<param name="movie" value="'+u+'">';
 flash_tag +='<param name="quality" value="high">';
 flash_tag +='<param name="wmode" value="transparent">';
 flash_tag +='<embed src="'+u+'" allowScriptAccess="sameDomain" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" ';
 flash_tag +='application/x-shockwave-flash"  width="'+w+'" height=="'+h+'"></embed></object>'
 document.write(flash_tag);
}


