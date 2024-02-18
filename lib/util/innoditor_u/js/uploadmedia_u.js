/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// uploadmedia.js
//					
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


var strLangPath = g_browserCHK.language;
document.write('<link rel="stylesheet" href="./css/' + strLangPath + '/uploadstyle.css" type="text/css">');
document.write('<script type="text/javascript" src="./js/res/' + strLangPath + '/res_upload_media.js"></scrip' +'t>');


var g_objUploadObjectProp = new Object();
g_objUploadObjectProp.media_path = "";
g_objUploadObjectProp.upload_url = "";
g_objUploadObjectProp.width = "300";
g_objUploadObjectProp.height = "300";
g_objUploadObjectProp.ShowControls = true;
g_objUploadObjectProp.ShowStatusBar = true;

var g_objLoadURLObjectProp = new Object();
g_objLoadURLObjectProp.media_url = "";
g_objLoadURLObjectProp.width = "300";
g_objLoadURLObjectProp.height = "300";
g_objLoadURLObjectProp.ShowControls = true;
g_objLoadURLObjectProp.ShowStatusBar = true;

var g_nSelectedTabIndex = 1;
var g_strImageDirPath = "./image/editwin/" + strLangPath + "/";

var res_image = new Array();
res_image[0] = "pop_upload_media_title";
res_image[1] = "btn_upload";
res_image[2] = "btn_url_load";
res_image[3] = "pop_upload_image_main_title_1";
res_image[4] = "btn_apply";
res_image[5] = "btn_cancel";
res_image[6] = "pop_upload_status";


function fnAdjustPage()
{
	try
	{
		fnInitResource();
	}
	catch(e)
	{
	}
}

function fnInitResource()
{
	var objTabImage_1 = document.getElementById("imgTab_1");
	objTabImage_1.src = g_strImageDirPath + "pop_upload_tab_media_1_r.gif";

	var objTabImage_2 = document.getElementById("imgTab_2");
	objTabImage_2.src = g_strImageDirPath + "pop_upload_tab_media_2.gif";

	for(var i=0; i<res_image.length; i++)
	{
		var objResImage = document.getElementById("res_image_" + i);

		var strImageSrc = g_strImageDirPath + res_image[i] + ".gif";
		objResImage.src = strImageSrc;
	}

	for(var i=0; i<res_item.length; i++)
	{
		var objResItem = document.getElementById("res_item_" + i);
		objResItem.innerHTML = res_item[i];
	}
}

function fnTabMouseOver(objTab, nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		objTab.src = g_strImageDirPath + "pop_upload_tab_media_" + nTabIndex + "_r.gif";
	}
}

function fnTabMouseOut(objTab, nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		objTab.src = g_strImageDirPath + "pop_upload_tab_media_" + nTabIndex + ".gif";
	}
}

function fnTabClick(nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		g_nSelectedTabIndex = nTabIndex;

		if(1 == nTabIndex)
		{
			document.body.all["imgTab_1"].src = g_strImageDirPath + "pop_upload_tab_media_1_r.gif";
			document.body.all["imgTab_2"].src = g_strImageDirPath + "pop_upload_tab_media_2.gif";

			document.body.all["mediaUploadArea"].style.display = "";
			document.body.all["mediaLoadURLArea"].style.display = "none";


			g_objLoadURLObjectProp.media_url = document.body.all["txtMediaPath"].value;
			g_objLoadURLObjectProp.width = document.body.all["txtWidth"].value;
			g_objLoadURLObjectProp.height = document.body.all["txtHeight"].value;
			g_objLoadURLObjectProp.ShowControls = (document.body.all["chkShowControls"].checked ? true : false);
			g_objLoadURLObjectProp.ShowStatusBar = (document.body.all["chkShowStatusBar"].checked ? true : false);

			document.body.all["txtMediaPath"].value = g_objUploadObjectProp.media_path;
			document.body.all["txtWidth"].value = g_objUploadObjectProp.width;
			document.body.all["txtHeight"].value = g_objUploadObjectProp.height;
			document.body.all["chkShowControls"].checked = g_objUploadObjectProp.ShowControls;
			document.body.all["chkShowStatusBar"].checked = g_objUploadObjectProp.ShowStatusBar;

			if("" == g_objUploadObjectProp.upload_url)
			{
				document.body.all["media_main"].innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
			}
			else
			{
				var strMediaURL = g_objUploadObjectProp.upload_url;
				fnLoadMedia(strMediaURL);
			}
		}
		else
		{
			document.body.all["imgTab_1"].src = g_strImageDirPath + "pop_upload_tab_media_1.gif";
			document.body.all["imgTab_2"].src = g_strImageDirPath + "pop_upload_tab_media_2_r.gif";

			document.body.all["mediaUploadArea"].style.display = "none";
			document.body.all["mediaLoadURLArea"].style.display = "";

			g_objUploadObjectProp.media_path = document.body.all["txtMediaPath"].value;
			g_objUploadObjectProp.width = document.body.all["txtWidth"].value;
			g_objUploadObjectProp.height = document.body.all["txtHeight"].value;
			g_objUploadObjectProp.ShowControls = (document.body.all["chkShowControls"].checked ? true : false);
			g_objUploadObjectProp.ShowStatusBar = (document.body.all["chkShowStatusBar"].checked ? true : false);

			document.body.all["txtMediaPath"].value = g_objLoadURLObjectProp.media_url;
			document.body.all["txtWidth"].value = g_objLoadURLObjectProp.width;
			document.body.all["txtHeight"].value = g_objLoadURLObjectProp.height;
			document.body.all["chkShowControls"].checked = g_objLoadURLObjectProp.ShowControls;
			document.body.all["chkShowStatusBar"].checked = g_objLoadURLObjectProp.ShowStatusBar;

			if("" == g_objLoadURLObjectProp.media_url)
			{
				document.body.all["media_main"].innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_2.gif' border='0'>";
			}
			else
			{
				var strMediaURL = g_objLoadURLObjectProp.media_url;

				var strCheckMediaURL = strMediaURL.toLowerCase();
				if( (-1 == strCheckMediaURL.indexOf(".wmv")) )
				{
				}
				else
				{
					fnLoadMedia(strMediaURL);
				}
			}
		}
	}
}

function fnLoadURLMedia()
{
	try
	{
		var strMediaURL = document.body.all["txtMediaPath"].value;
		if("" == strMediaURL)
		{
			alert(res_warning_msg[0]);
			document.body.all["txtMediaPath"].focus();
			return;
		}

		g_objLoadURLObjectProp.media_url = strMediaURL;
		fnLoadMedia(strMediaURL);
	}
	catch(e)
	{
	}
}

function fnLoadMedia(strMediaURL)
{
	try
	{
		if(("" == strMediaURL) || ("undefined" == strMediaURL)) return;

		var nWidth = document.body.all["txtWidth"].value;
		var nHeight = document.body.all["txtHeight"].value;
		var bShowControls = document.body.all["chkShowControls"].checked;
		var bShowStatusBar = document.body.all["chkShowStatusBar"].checked;

		var strMediaHTML = fnMakeMediaHTML(strMediaURL, nWidth, nHeight, bShowControls, bShowStatusBar);
		document.body.all["media_main"].innerHTML = strMediaHTML;
	}
	catch(e)
	{
	}
}

function fnMakeMediaHTML(strMediaURL, nWidth, nHeight, bShowControls, bShowStatusBar)
{
	try
	{
		var strShowControls = (bShowControls ? "true" : "false");
		var strShowStatusBar = (bShowStatusBar ? "true" : "false");

		var strHTML = '';
		strHTML += '<EMBED src="' + strMediaURL + '" ';
		strHTML +=		'width=' + nWidth + ' height=' + nHeight + '" ';
		strHTML +=		'autostart="true" ';
		strHTML +=		'ShowControls="' + strShowControls + '" ';
		strHTML +=		'ShowStatusBar="' + strShowStatusBar + '" ';
//		strHTML +=		'type="video/x-ms-wmv" />';
		strHTML +=		' />';

		return strHTML;
	}
	catch(e)
	{
	}

	return "";
}

function fnChangeFile(strMediaPath)
{
	var strUploadMediaPath = strMediaPath.toLowerCase();

	if( (-1 == strMediaPath.indexOf(".wmv")) )
	{
		alert(res_warning_msg[1]);

		g_objUploadObjectProp.media_path = "";
		g_objUploadObjectProp.upload_url = "";
	}
	else
	{
		g_objUploadObjectProp.media_path = strMediaPath;
		g_objUploadObjectProp.upload_url = "";
	}

	document.body.all["txtMediaPath"].value = g_objUploadObjectProp.media_path;
	document.body.all["media_main"].innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
}

function fnUploadFile()
{
	var objFrmUpload = document.frmUpload;
	var objFile = objFrmUpload.fileUpload;

	var strMediaPath = objFile.value;
	strMediaPath = strMediaPath.toLowerCase();

	if(0 == strMediaPath.length || "" == strMediaPath)
	{
		alert(res_warning_msg[2]);
		return;
	}

	if( (-1 == strMediaPath.indexOf(".wmv")) )
	{
		alert(res_warning_msg[1]);
		return;
	}

	document.body.all["uploadStatus"].style.display = "";

	objFrmUpload.action = g_strUploadPath_Media;
	objFrmUpload.target = "ifrmUpload";
	objFrmUpload.submit();
}

function fnUploadResult(strUploadMediaURL)
{
	if("" == strUploadMediaURL)
	{
		alert(res_warning_msg[3]);

		g_objUploadObjectProp.media_path = "";
		g_objUploadObjectProp.upload_url = "";
	}
	else
	{
		g_objUploadObjectProp.upload_url = strUploadMediaURL;

		fnLoadMedia(strUploadMediaURL);
	}

	document.body.all["uploadStatus"].style.display = "none";
}
