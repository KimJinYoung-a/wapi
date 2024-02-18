var nSelectedItemIndex = 0;
var nItemLength = 0;
var strItemClassName = "";

function fnHelpItemOver(objItem, nItemIndex)
{
	if(nSelectedItemIndex != nItemIndex)
	{
		objItem.className = strItemClassName + "_r";
	}
}

function fnHelpItemOut(objItem, nItemIndex)
{
	if(nSelectedItemIndex != nItemIndex)
	{
		objItem.className = strItemClassName;
	}
}

function fnHelpItemClick(nItemIndex)
{
	if(nSelectedItemIndex != nItemIndex)
	{
		nSelectedItemIndex = nItemIndex;

		for(var i=0; i<nItemLength; i++)
		{
			var objItem = document.getElementById("helpItem_" + i);
			var objItemMain = document.getElementById("helpItemMain_" + i);

			if(i != nItemIndex)
			{
				objItem.className = strItemClassName;
				objItemMain.style.display = "none";
			}
			else
			{
				objItemMain.style.display = "";
			}
		}
	}
}


var g_strLangPath = g_browserCHK.language;
document.write('<script type="text/javascript" src="./js/' + g_strLangPath + '/help_resource_u.js"></scrip' +'t>');


function fnSetHelpResourceDetail(nDetailType)
{
	for(var i=0; i<nItemLength; i++)
	{
		var objItem = document.getElementById("helpItem_" + i);

		var objItemMainTitle = document.getElementById("helpItemMainTitle_" + i);
		var objItemMainDescription = document.getElementById("helpItemMainDescription_" + i);
		var objItemMainImage = document.getElementById("helpItemMainImage_" + i);


		objItem.innerHTML = res_helpitem[nDetailType][i];

		if(1 == nDetailType)
		{
			objItemMainTitle.innerHTML = "<b>" + res_helpitem[nDetailType][i] + res_helpitem_MainTitle_Sub[i] + "</b>";
		}
		else
		{
			objItemMainTitle.innerHTML = "<b>" + res_helpitem[nDetailType][i] + "</b>";
		}

		objItemMainDescription.innerHTML = res_helpitem_MainDescription[nDetailType][i];

		var strImageSrc = "../help/image/" + g_strLangPath + "/help_main_" + nDetailType + "_" + i + ".gif";
		objItemMainImage.src = strImageSrc;
	}

	if(1 == nDetailType)
	{
		for(var i=0; i<res_helpitem_SubDescription_1.length; i++)
		{
			var objItemSubDescription = document.getElementById("helpItemSubDescription_" + i);
			objItemSubDescription.innerHTML = res_helpitem_SubDescription_1[i];
		}
	}
}
