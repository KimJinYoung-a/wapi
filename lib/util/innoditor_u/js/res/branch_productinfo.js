/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// branch_productinfo.js
//						
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

document.write('<script type="text/javascript" src="./js/res/' + g_browserCHK.language + '/res_productinfo.js"></scrip' +'t>');

function fnInitResource()
{
	for(var i=0; i<res_item.length; i++)
	{
		var objResItem = document.getElementById("res_item_" + i);
		objResItem.innerHTML = res_item[i];
	}

	var objButton = document.getElementById("btnConfirm");
	objButton.value = res_button_text;
}
