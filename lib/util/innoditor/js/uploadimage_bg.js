/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// uploadImage_bg.js -	배경 image 파일 Upload 및 설정 관련
//					
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


// 변수 선언
var g_objUploadObject = null;
var g_strUploadImageURL = "";

var g_strImageDirPath = "./image/editwin/ko/";


function fnAdjustPage()
{
	try
	{
		var objFile = document.getElementById("fileUpload");

		if(g_browserCHK.ie)
		{
			objFile.onpropertychange = fnChangeImage;
		}
		else if(g_browserCHK.ff)
		{
			objFile.onchange = fnChangeImage;
			objFile.style.left = -160;
		}
		else if(g_browserCHK.wk || g_browserCHK.op)
		{
			objFile.onchange = fnChangeImage;
		}
	}
	catch(e)
	{
	}
}

function fnChangeImage()
{
	try
	{
		var objFileValue = document.getElementById("fileUpload").value;
		fnLoadImage(objFileValue, 1, true, true);
	}
	catch(e)
	{
	}
}

function fnInitLoadImage(strImageURL)
{
	try
	{
		var strCopyImageURL = strImageURL;
		strCopyImageURL = strCopyImageURL.toLowerCase();

		if(("" == strCopyImageURL) || ("none" == strCopyImageURL)) return;

		fnLoadImage(strImageURL, 2);
	}
	catch(e)
	{
	}
}

function fnInitLoadRepeat(strRepeat)
{
	try
	{
		strRepeat = strRepeat.toLowerCase();
		var objSelRepeat = document.getElementById("selBackgroundRepeat");
		for(var i=0; objSelRepeat.length; i++)
		{
			if(strRepeat == objSelRepeat.options[i].value)
			{
				objSelRepeat.options[i].selected = true;
				break;
			}
		}
	}
	catch(e)
	{
	}
}

function fnLoadImage(strUploadImageURL, nLoadType)
{
	try
	{
		g_strUploadImageURL = "";
		if(("" == strUploadImageURL) || ("undefined" == strUploadImageURL)) return;

		g_strUploadImageURL = strUploadImageURL;

		if(1 == nLoadType)
		{
			if(g_browserCHK.ie)
			{
				if(g_browserCHK.ie6)
				{
					setTimeout("fnLoadImageInfo(" + nLoadType + ")", 100);
				}
				else
				{
					var nFilePathLength = strUploadImageURL.length;
					var nFileNameIndex = strUploadImageURL.lastIndexOf("\\");
					var strFileName = strUploadImageURL.substr(nFileNameIndex+1, nFilePathLength-nFileNameIndex);

					var strFileInfoHTML = fnFileInfoHTML(strFileName);

					document.getElementById("img_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
					document.getElementById("img_preview").innerHTML = strFileInfoHTML;
				}
			}
			else if(g_browserCHK.ff || g_browserCHK.wk || g_browserCHK.op)
			{
				var nFilePathLength = strUploadImageURL.length;
				var nFileNameIndex = strUploadImageURL.lastIndexOf("\\");
				var strFileName = strUploadImageURL.substr(nFileNameIndex+1, nFilePathLength-nFileNameIndex);

				var strFileInfoHTML = fnFileInfoHTML(strFileName);

				document.getElementById("img_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";
				document.getElementById("img_preview").innerHTML = strFileInfoHTML;
			}
		}
		else
		{
			setTimeout("fnLoadImageInfo(" + nLoadType + ")", 100);
		}
	}
	catch(e)
	{
	}
}

function fnFileInfoHTML(strFileName)
{
	try
	{
		var strFileInfoHTML = "";
		strFileInfoHTML += "<table align='center' width='144' height='50' style='table-layout:fixed;' border='0' cellpadding='0' cellspacing='0'>";
		strFileInfoHTML += "<tr>";
		strFileInfoHTML +=		"<td align='center' height='20' class='textDisplay'>선택한 파일명</td>";
		strFileInfoHTML += "</tr>";
		strFileInfoHTML += "<tr>";
		strFileInfoHTML +=		"<td align='center' height='30' class='textDisplay'><b>" + strFileName + "</b></td>";
		strFileInfoHTML += "</tr>";
		strFileInfoHTML += "</table>";

		return strFileInfoHTML;
	}
	catch(e)
	{
	}

	return "";
}

function fnLoadImageInfo(nLoadType)
{
	try
	{
		g_objUploadObject = null;
		g_objUploadObject = new Image();
		g_objUploadObject.src = g_strUploadImageURL;

		fnLoadImageSize(nLoadType);
	}
	catch(e)
	{
	}
}

function fnLoadImageSize(nLoadType)
{
	try
	{
		var nWidth = g_objUploadObject.width;
		var nHeight = g_objUploadObject.height;

		if(0 == nWidth || 0 == nHeight)
		{
			setTimeout("fnLoadImageSize(" + nLoadType + ")", 100);
		}
		else
		{
			fnLoadImgSizeDone(nLoadType);
		}
	}
	catch(e)
	{
	}
}

function fnLoadImgSizeDone(nLoadType)
{
	try
	{
		var nWidth = g_objUploadObject.width;
		var nHeight = g_objUploadObject.height;

		var nCheckWidth = nWidth;
		var nCheckHeight = nHeight;

		if(1 == nLoadType)
		{
			if(nWidth > 146) nCheckWidth = 146;
			if(nHeight > 146) nCheckHeight = 146;
		}
		else if(2 == nLoadType)
		{
			if(nWidth > 396) nCheckWidth = 396;
			if(nHeight > 406) nCheckHeight = 406;
		}
		else return;

		var strImageHTML = fnMakeImageHTML(nCheckWidth, nCheckHeight);

		if(1 == nLoadType)
		{
			var objPreview = document.getElementById("img_preview");
			objPreview.innerHTML = strImageHTML;

			document.getElementById("img_main").innerHTML = "<img src='" + g_strImageDirPath + "pop_upload_image_main_title_1.gif' border='0'>";

			g_objUploadObject = null;
		}
		else if(2 == nLoadType)
		{
			var objImgMain = document.getElementById("img_main");
			objImgMain.innerHTML = strImageHTML;
		}
		else return;
	}
	catch(e)
	{
	}
}

function fnMakeImageHTML(nWidth, nHeight)
{
	try
	{
		var strImageHTML = "";

		strImageHTML = "<img src='" + g_objUploadObject.src + "' ";
		strImageHTML += "width='" + nWidth + "' ";
		strImageHTML += "height='" + nHeight + "' border='0'>";

		return strImageHTML;
	}
	catch(e)
	{
	}

	return "";
}

function fnUploadFile()
{
	try
	{
		var objFrmUpload = document.frmUpload;
		var objFile = objFrmUpload.fileUpload;

		var strImageFilePath = objFile.value;
		strImageFilePath = strImageFilePath.toLowerCase();

		if(0 == strImageFilePath.length || "" == strImageFilePath)
		{
			g_strUploadImageURL = "";
			alert("파일선택에서 이미지파일을 먼저 선택하세요");
			return;
		}

		if((-1==strImageFilePath.indexOf(".jpg")) && (-1==strImageFilePath.indexOf(".gif")) && (-1==strImageFilePath.indexOf(".png")))
		{
			alert("이미지파일(jpg,gif,png) 파일만 올리실 수 있습니다");
			return;
		}


		// 파일 확장자 추가검사 - Start
		var nFindIndex = -1;
		var nFilePathLength = strImageFilePath.length;
		if(-1 != strImageFilePath.indexOf(".jpg"))
		{
			nFindIndex = strImageFilePath.indexOf(".jpg");
			if(nFilePathLength != (nFindIndex+4))
			{
				alert("이미지파일(jpg,gif,png) 파일만 올리실 수 있습니다");
				return;
			}
		}
		if(-1 != strImageFilePath.indexOf(".gif"))
		{
			nFindIndex = strImageFilePath.indexOf(".gif");
			if(nFilePathLength != (nFindIndex+4))
			{
				alert("이미지파일(jpg,gif,png) 파일만 올리실 수 있습니다");
				return;
			}
		}
		if(-1 != strImageFilePath.indexOf(".png"))
		{
			nFindIndex = strImageFilePath.indexOf(".png");
			if(nFilePathLength != (nFindIndex+4))
			{
				alert("이미지파일(jpg,gif,png) 파일만 올리실 수 있습니다");
				return;
			}
		}
		// 파일 확장자 추가검사 - End


		document.getElementById("uploadStatus").style.display = "";

		objFrmUpload.action = g_strUploadPath_Image;
		objFrmUpload.target = "ifrmUpload";
		objFrmUpload.submit();
	}
	catch(e)
	{
	}
}

function fnUploadResult(strUploadImageURL)
{
	if("" == strUploadImageURL)
	{
		alert("업로드가 실패하였습니다\n이미지 파일이 아니거나 용량을 초과하였습니다");
		g_strUploadImageURL = "";
	}
	else
	{
		fnLoadImage(strUploadImageURL, 2);
	}

	document.getElementById("uploadStatus").style.display = "none";
}
