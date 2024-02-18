/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// branch_doc_templete.js
//						
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

var strLangPath = g_browserCHK.language;
document.write('<link rel="stylesheet" href="./css/' + strLangPath + '/uploadstyle.css" type="text/css">');

var g_strImageDirPath = "./image/editwin/" + strLangPath + "/";

var res_image = new Array();
res_image[0] = "pop_doc_templete_title";
res_image[1] = "pop_doc_templete_tab_image_1_r";
res_image[2] = "btn_cancel";


function fnInitResource()
{
	for(var i=0; i<res_image.length; i++)
	{
		var objResImage = document.getElementById("res_image_" + i);

		var strImageSrc = g_strImageDirPath + res_image[i] + ".gif";
		objResImage.src = strImageSrc;
	}
}
