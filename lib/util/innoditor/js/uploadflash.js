/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// uploadflash.js -	flash 파일 Upload 관련
//					
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// 변수 선언
var g_objUploadObjectProp = new Object();
g_objUploadObjectProp.flash_path = "";
g_objUploadObjectProp.upload_url = "";
g_objUploadObjectProp.width = "300";
g_objUploadObjectProp.height = "300";
g_objUploadObjectProp.align = "";
g_objUploadObjectProp.border = "0";
g_objUploadObjectProp.hspace = "0";
g_objUploadObjectProp.vspace = "0";

var g_objLoadURLObjectProp = new Object();
g_objLoadURLObjectProp.flash_url = "http://www.innoditor.com/sample/flash/sample_4.swf";
g_objLoadURLObjectProp.width = "300";
g_objLoadURLObjectProp.height = "300";
g_objLoadURLObjectProp.align = "";
g_objLoadURLObjectProp.border = "0";
g_objLoadURLObjectProp.hspace = "0";
g_objLoadURLObjectProp.vspace = "0";

var g_nSelectedTabIndex = 1;

var g_strImageDirPath = "./image/editwin/ko/";


function fnAdjustPage()
{
	try
	{
		if(g_browserCHK.ff)
		{
			var objFile = document.getElementById("fileUpload");
			objFile.style.left = -160;
		}
	}
	catch(e)
	{
	}
}

function fnTabMouseOver(objTab, nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		objTab.src = g_strImageDirPath + "pop_upload_tab_flash_" + nTabIndex + "_r.gif";
	}
}

function fnTabMouseOut(objTab, nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		objTab.src = g_strImageDirPath + "pop_upload_tab_flash_" + nTabIndex + ".gif";
	}
}

function fnTabClick(nTabIndex)
{
	if(nTabIndex != g_nSelectedTabIndex)
	{
		g_nSelectedTabIndex = nTabIndex;

		if(1 == nTabIndex)
		{
			document.getElementById("imgTab_1").src = g_strImageDirPath + "pop_upload_tab_flash_1_r.gif";
			document.getElementById("imgTab_2").src = g_strImageDirPath + "pop_upload_tab_flash_2.gif";

			document.getElementById("flashUploadArea").style.display = "";
			document.getElementById("flashLoadURLArea").style.display = "none";


			g_objLoadURLObjectProp.flash_url = document.getElementById("txtFlashPath").value;
			g_objLoadURLObjectProp.width = document.getElementById("txtWidth").value;
			g_objLoadURLObjectProp.height = document.getElementById("txtHeight").value;
			g_objLoadURLObjectProp.align = document.getElementById("selAlign").value;
			g_objLoadURLObjectProp.border = document.getElementById("txtBorder").value;
			g_objLoadURLObjectProp.hspace = document.getElementById("txtHSpace").value;
			g_objLoadURLObjectProp.vspace = document.getElementById("txtVSpace").value;

			document.getElementById("txtFlashPath").value = g_objUploadObjectProp.flash_path;
			document.getElementById("txtWidth").value = g_objUploadObjectProp.width;
			document.getElementById("txtHeight").value = g_objUploadObjectProp.height;
			document.getElementById("selAlign").value = g_objUploadObjectProp.align;
			document.getElementById("txtBorder").value = g_objUploadObjectProp.border;
			document.getElementById("txtHSpace").value = g_objUploadObjectProp.hspace;
			document.getElementById("txtVSpace").value = g_objUploadObjectProp.vspace;

			if("" == g_objUploadObjectProp.upload_url)
			{
				document.getElementById("flash_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
			}
			else
			{
				var strFlashURL = g_objUploadObjectProp.upload_url;
				fnLoadFlash(strFlashURL);
			}
		}
		else
		{
			document.getElementById("imgTab_1").src = g_strImageDirPath + "pop_upload_tab_flash_1.gif";
			document.getElementById("imgTab_2").src = g_strImageDirPath + "pop_upload_tab_flash_2_r.gif";

			document.getElementById("flashUploadArea").style.display = "none";
			document.getElementById("flashLoadURLArea").style.display = "";

			g_objUploadObjectProp.flash_path = document.getElementById("txtFlashPath").value;
			g_objUploadObjectProp.width = document.getElementById("txtWidth").value;
			g_objUploadObjectProp.height = document.getElementById("txtHeight").value;
			g_objUploadObjectProp.align = document.getElementById("selAlign").value;
			g_objUploadObjectProp.border = document.getElementById("txtBorder").value;
			g_objUploadObjectProp.hspace = document.getElementById("txtHSpace").value;
			g_objUploadObjectProp.vspace = document.getElementById("txtVSpace").value;

			document.getElementById("txtFlashPath").value = g_objLoadURLObjectProp.flash_url;
			document.getElementById("txtWidth").value = g_objLoadURLObjectProp.width;
			document.getElementById("txtHeight").value = g_objLoadURLObjectProp.height;
			document.getElementById("selAlign").value = g_objLoadURLObjectProp.align;
			document.getElementById("txtBorder").value = g_objLoadURLObjectProp.border;
			document.getElementById("txtHSpace").value = g_objLoadURLObjectProp.hspace;
			document.getElementById("txtVSpace").value = g_objLoadURLObjectProp.vspace;

			if("" == g_objLoadURLObjectProp.flash_url)
			{
				document.getElementById("flash_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_2.gif' border='0'>";
			}
			else
			{
				var strFlashURL = g_objLoadURLObjectProp.flash_url;
				fnLoadFlash(strFlashURL);
			}
		}
	}
}

function fnLoadURLFlash()
{
	try
	{
		var strFlashURL = document.getElementById("txtFlashPath").value;
		if("" == strFlashURL)
		{
			alert("플래시(swf) URL을 입력하세요");
			document.getElementById("txtFlashPath").focus();
			return;
		}

		g_objLoadURLObjectProp.flash_url = strFlashURL;
		fnLoadFlash(strFlashURL);
	}
	catch(e)
	{
	}
}

function fnLoadFlash(strFlashURL)
{
	try
	{
		if(("" == strFlashURL) || ("undefined" == strFlashURL)) return;

		var nWidth = document.getElementById("txtWidth").value;
		var nHeight = document.getElementById("txtHeight").value;
		var strAlign = document.getElementById("selAlign").value;
		var nBorder = document.getElementById("txtBorder").value;
		var nHSpace = document.getElementById("txtHSpace").value;
		var nVSpace = document.getElementById("txtVSpace").value;

		if(nWidth > 390) nWidth = 390;
		if(nHeight > 400) nHeight = 400;

		var strFlashHTML = fnMakeFlashHTML(strFlashURL, nWidth, nHeight, strAlign, nBorder, nHSpace, nVSpace);
		document.getElementById("flash_main").innerHTML = strFlashHTML;
	}
	catch(e)
	{
	}
}

function fnMakeFlashHTML(strFlashURL, nWidth, nHeight, strAlign, nBorder, nHSpace, nVSpace)
{
	try
	{
		var strHTML = '';
		strHTML += '<object width="' + nWidth + '" height="' + nHeight + '" align="' + strAlign + '" border="' + nBorder + '" hspace="' + nHSpace + '" vspace="' + nVSpace + '" ';
		strHTML +=		'CLASSID="CLSID:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,47,0">';
		strHTML += '<param name="movie" value="' + strFlashURL + '">';
		strHTML += '<param name="quality" value="high">';
		strHTML += '<param name="allowFullScreen" value="true">';
		strHTML += '<embed src="' + strFlashURL + '" ';
		strHTML +=		'quality="high" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />';
		strHTML += '</object>';

		return strHTML;
	}
	catch(e)
	{
	}

	return "";
}

function fnChangeFile(strFlashPath)
{
	var strUploadFlashPath = strFlashPath.toLowerCase();

	if(-1 == strFlashPath.indexOf(".swf"))
	{
		alert("플래시 파일(swf)만 올리실 수 있습니다");

		g_objUploadObjectProp.flash_path = "";
		g_objUploadObjectProp.upload_url = "";
	}
	else
	{
		g_objUploadObjectProp.flash_path = strFlashPath;
		g_objUploadObjectProp.upload_url = "";
	}

	document.getElementById("txtFlashPath").value = g_objUploadObjectProp.flash_path;
	document.getElementById("flash_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
}

function fnUploadFile()
{
	try
	{
		var objFrmUpload = document.frmUpload;
		var objFile = objFrmUpload.fileUpload;

		var strFlashPath = objFile.value;
		strFlashPath = strFlashPath.toLowerCase();

		if(0 == strFlashPath.length || "" == strFlashPath)
		{
			alert("파일선택에서 플래시 파일을 먼저 선택하세요");
			return;
		}

		if(-1 == strFlashPath.indexOf(".swf"))
		{
			alert("플래시 파일(swf)만 올리실 수 있습니다");
			return;
		}


		// 파일 확장자 추가검사 - Start
		var nFindIndex = -1;
		var nFilePathLength = strFlashPath.length;
		if(-1 != strFlashPath.indexOf(".swf"))
		{
			nFindIndex = strFlashPath.indexOf(".swf");
			if(nFilePathLength != (nFindIndex+4))
			{
				alert("플래시 파일(swf)만 올리실 수 있습니다");
				return;
			}
		}
		// 파일 확장자 추가검사 - End


		document.getElementById("uploadStatus").style.display = "";

		objFrmUpload.action = g_strUploadPath_Flash;
		objFrmUpload.target = "ifrmUpload";
		objFrmUpload.submit();
	}
	catch(e)
	{
	}
}

function fnUploadResult(strUploadFlashURL)
{
	if("" == strUploadFlashURL)
	{
		alert("업로드가 실패하였습니다\n플래시 파일이 아니거나 용량을 초과하였습니다");

		g_objUploadObjectProp.flash_path = "";
		g_objUploadObjectProp.upload_url = "";
	}
	else
	{
		g_objUploadObjectProp.upload_url = strUploadFlashURL;

		fnLoadFlash(strUploadFlashURL);
	}

	document.getElementById("uploadStatus").style.display = "none";
}

