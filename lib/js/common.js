function jsPopCalendar(fName,sName){
	var fd = eval("document."+fName+"."+sName);

    //calendar.js 포함 되어있으면
    if (typeof calPopup == "function"){
        var compname = 'document.' + fName + '.' + sName;
        calPopup(fd,'calendarPopup',20+80,0, compname,'');
    }else{
    	var winCal;
    	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}

function PrintBarcodeOfflineOrder(masteridx) {

	var popwin = window.open("/common/popBarcodePrint.asp?masteridx=" + masteridx + "&barcodetype=offlineorder","PrintBarcodeOfflineOrder","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PrintBarcodeOfflineOrderByBox(masteridx) {

	var popwin = window.open("/common/popBarcodePrint.asp?masteridx=" + masteridx + "&barcodetype=offlineorderbox","PrintBarcodeOfflineOrderByBox","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PrintBarcodeOfflineOrderByOneBox(masteridx, boxno) {

	var popwin = window.open("/common/popBarcodePrint.asp?masteridx=" + masteridx + "&boxno=" + boxno + "&barcodetype=offlineorderbox","PrintBarcodeOfflineOrderByOneBox","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ViewOfflineOrderSheet(masteridx) {

	var popwin = window.open("/common/popOrderSheetView.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ViewOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ExcelOfflineOrderSheet(masteridx) {

	var popwin = window.open("/common/popOrderSheetExcel.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ExcelOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//범용버전 		//20120831 한용민 생성
function ExcelOfflineOrderSheetpublic(masteridx) {

	var ExcelOfflineOrderSheetpublic = window.open("/common/popOrderpublicSheetExcel.asp?masteridx=" + masteridx + "&ordersheettype=offlineorder","ExcelOfflineOrderSheetpublic","width=800 height=600 scrollbars=yes resizable=yes");
	ExcelOfflineOrderSheetpublic.focus();
}

function fnTrim(orgStr){
    return orgStr.replace(/(^\s*)|(\s*$)/gi, "");
}

function phone_format(obj) {
	var tmp;

	tmp = obj.value;

	tmp = tmp.replace(/\-/g, "");

	if (isNaN(tmp) == true) {
		alert("전화번호에는 숫자이외에 입력할 수 없습니다.");
		obj.value = "";
		obj.focus();
		return;
	}

	if (tmp.length <= 4) {
		obj.value = tmp;
	} else if (tmp.length <= 8) {
		obj.value = tmp.replace(/([0-9]+)([0-9]{4})/,"$1-$2");
	} else {
		obj.value = tmp.replace(/(^02.{0}|^01.{1}|^070|[0-9]{3})([0-9]+)([0-9]{4})/,"$1-$2-$3");
	}
}

function fnCheckAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        comp.checked = bool;
        AnCheckClick(comp);
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
            AnCheckClick(comp[i]);
        }
    }
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

//상품 배송소요일 LIST
function popItemAvgDlvList(itemid){
    //return;
    var popwin = window.open("/admin/datamart/baesong/iframe_baesong_term_list.asp?itemid=" + itemid,"popItemAvgDlvList","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//상품 배송소요일 graph
function popItemAvgDlvGraph(itemid){
    //return;
    var popwin = window.open("/admin/datamart/baesong/iframe_baesong_term_graph.asp?itemid=" + itemid,"popItemAvgDlvGraph","width=900 height=500 scrollbars=yes resizable=yes");
	popwin.focus();
}

//브랜드 ID 검색 팝업창
function jsSearchBrandID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/lib/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandIDwithUserDIV(frmName,compName, userdiv) {
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&userdiv=" + userdiv,"jsSearchBrandIDwithUserDIV","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchMeachulID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popMeachulIDSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"jsSearchMeachulID","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandID2(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch2.asp?frmName=" + frmName + "&compName=" + compName + "&rect=","popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsSearchBrandIDchgMargin(frmName,compName,mgnfName,evtjs){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch_chgMargin.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal + "&mgnfName=" + mgnfName + "&evtjs=" + evtjs,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function PopBrandInfoEdit(makerid){
    var popwin = window.open("/admin/member/popbrandinfoonly.asp?designer=" + makerid,"popbrandinfoonly","width=660 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheInfoEdit(groupid){
    var popwin = window.open("/admin/member/popupcheinfoonly.asp?groupid=" + groupid,"popupcheinfoonly","width=660 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheBrandInfoEdit(makerid){
    var popwin = window.open("/admin/member/popupchebrandinfo.asp?designer=" + makerid,"popupchebrandinfoedit","width=660 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopBrandAdminUsingChange(v){
	var popwin = window.open("/admin/member/popbrandadminusing.asp?designer=" + v,"popbrandadminusing","width=640 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheSelect(frmname){
	var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname,"popupcheselect","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheSelect_shop(frmname,shopyn){
	var PopUpcheSelect_shop = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&shopyn="+shopyn,"PopUpcheSelect_shop","width=800 height=580 scrollbars=yes resizable=yes");
	PopUpcheSelect_shop.focus();
}

function TnPopItemStock(itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopItemIpChulList(fromdate,todate,itemgubun,itemid,itemoption,ipchulflag){
	var popwin = window.open('/common/pop_stock_ipgo.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag, 'pop_stock_ipgo', 'width=800,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function PopItemIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption,ipchulflag, shopid){
	var popwin = window.open('/common/pop_ipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid, 'pop_stock_ipgo_off', 'width=1000,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function calendarOpen4(objTarget,caption,defaultval){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.asp?caption=" + caption + "&defaultval=" + defaultval , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	  	return true;
	 }else{
		return false;
	 }
}


function calendarOpen3(objTarget,caption,defaultval){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.asp?caption=" + caption + "&defaultval=" + defaultval , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	 }else{

	 }
}

function calendarOpen2(objTarget){
	 //var objTarget = document.getElementById(targetName);
	 output = window.showModalDialog("/lib/calendar.html" , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	  	return true;
	 }else{
	 	return false;
	 }
}

function calendarOpen(objTarget){
/*
	 output = window.showModalDialog("/lib/calendar.html" , null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 if(output!=''){
	  	objTarget.value=output;
	 }else{

	 }
*/
    //calendar.js 포함 되어있으면
    if (typeof calPopup == "function"){
        var compname = 'document.' + objTarget.form.name + '.' + objTarget.name;
        calPopup(objTarget,'calendarPopup',20+80,0, compname,'');
    }else{
        var fName = objTarget.form.name;
        var sName = objTarget.name;
    	var winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}

function AnModiNormal(frm){
	var ret = confirm('수정하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function AnDelNormal(frm){
	var ret = confirm('삭제 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function AnPopItemList(){
	window.open('/module/popitemlistbydesigner.asp', 'itemlist', 'width=600,height=660,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
}

function CheckNDolimitsell(frm){
	if (!IsDigit(frm.limitno.value)){
		alert('한정수량은 숫자만 가능합니다.');
		frm.limitno.focus();
		return false;
	}

	if (!IsDigit(frm.limitsold.value)){
		alert('판매된 수량은 숫자만 가능합니다.');
		frm.limitsold.focus();
		return false;
	}

	if (frm.baedalcd.value.length<1){
		alert('배송구분을 선택하세요..');
		frm.baedalcd.focus();
		return false;
	}

	var ret = confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function CheckNDobuyprice(frm){
	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return false;
	}

	if (!IsDigit(frm.sellvat.value)){
		alert('판매Vat는 숫자만 가능합니다.');
		frm.sellvat.focus();
		return false;
	}

	if (!IsDigit(frm.buycash.value)){
		alert('구매가는 숫자만 가능합니다.');
		frm.buycash.focus();
		return false;
	}

	if (!IsDigit(frm.buyvat.value)){
		alert('구매Vat는 숫자만 가능합니다.');
		frm.buyvat.focus();
		return false;
	}

	if (!IsDigit(frm.marginrate.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.marginrate.focus();
		return false;
	}

	if ((frm.marginrate.value<0)&&(frm.marginrate.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.marginrate.focus();
		return false;
	}
	var ret = confirm('저장하시겠습니까?');

	return ret;
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}

function hL(E){
	while (E.tagName!="TR")
	{
		E=E.parentElement;
	}

	E.className = "H";
}

function dL(E){
	while (E.tagName!="TR"){
		E=E.parentElement;
	}

	E.className = "";
}

function IsInteger(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("-0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

function IsDigit(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

function IsDouble(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789.".indexOf(v.charAt(j)) < 0) {
			return false;
		}
	}
	return true;
}

function AnCheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문을 새 발주서로 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.submit();
	}
}

// 오프라인발주서작성
function AnCheckNBaljuOffLine(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문을 새 발주서로 저장하시겠습니까?');
	if (ret){
		upfrm.orderidx.value = "";
		upfrm.baljucode.value = "";
		upfrm.baljuid.value = "";

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderidx.value = upfrm.orderidx.value + "|" + frm.orderidx.value;
					upfrm.baljucode.value = upfrm.baljucode.value + "|" + frm.baljucode.value;
					upfrm.baljuid.value = upfrm.baljuid.value + "|" + frm.baljuid.value;
				}
			}
		}
		upfrm.submit();
		// alert(upfrm.baljuid.value);
	}
}

function AnAllCalcu(frm){
	var frmtarget;

	if (!IsDouble(frm.mgall.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.mgall.focus();
		return;
	}

	if ((frm.mgall.value<0)&&(frm.mgall.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.mgall.focus();
		return;
	}

	//if ((!frm.rdall[0].checked)&&(!frm.rdall[1].checked)){
	//	alert('부가세 포함여부를 선택하세요.');
	//	return;
	//}

	for (var i=0;i<document.forms.length;i++){
		frmtarget = document.forms[i];
		if (frmtarget.name.substr(0,9)=="frmBuyPrc") {
			if (!frmtarget.cksel.checked) continue;
			frmtarget.marginrate.value = frm.mgall.value;
			//if (frm.rdall[0].checked){
			//	frmtarget.vtinclude[0].checked=true;
				AnAutoCalcu(frmtarget,true);
			//}else{
			//	frmtarget.vtinclude[1].checked=true;
			//	AnAutoCalcu(frmtarget,false);
			//}
		}
	}
}

function AnAutoCalcu(frm,bool){
	var buf;
	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(frm.marginrate.value)){
		alert('마진율은 숫자만 가능합니다.');
		frm.marginrate.focus();
		return;
	}

	if ((frm.marginrate.value<0)&&(frm.marginrate.value>100)){
		alert('마진율은 0~100만 가능합니다.');
		frm.marginrate.focus();
		return;
	}

	if (bool){
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		buf = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100.0)));
		frm.buycash.value = buf;
		frm.buyvat.value = parseInt(Math.round(buf/11));
		frm.tmpbuycash.value = parseInt(Math.round(buf-frm.buyvat.value));
		//frm.buyvat.value = Math.floor(buf/11);
		//frm.tmpbuycash.value = Math.floor(buf-frm.buyvat.value);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1 + frm.tmpbuycash.value * 1);

	}else{
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		frm.tmpbuycash.value = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100)));
		frm.buycash.value = parseInt(Math.round(frm.tmpbuycash.value*1.1));
		frm.buyvat.value = parseInt(Math.round(frm.tmpbuycash.value*0.1));
		//frm.tmpbuycash.value = Math.floor(frm.sellcash.value*(1-frm.marginrate.value/100));
		//frm.buyvat.value = Math.floor(frm.buycash.value/10);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1  + frm.tmpbuycash.value * 1 );
	}
}

function AnDesignerSearchSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 업체가 없습니다.');
		return;
	}

	var ret = confirm('선택 업체를 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.designerid.value = upfrm.designerid.value + "|" + frm.designerid.value;

					if (frm.isusing[0].checked){
						upfrm.isusing.value = upfrm.isusing.value + "|" + "Y";
					}else{
						upfrm.isusing.value = upfrm.isusing.value + "|" + "N";
					}

					if (frm.isextusing[0].checked){
						upfrm.isextusing.value = upfrm.isextusing.value + "|" + "Y";
					}else{
						upfrm.isextusing.value = upfrm.isextusing.value + "|" + "N";
					}

					if (frm.isb2b[0].checked){
						upfrm.isb2b.value = upfrm.isb2b.value + "|" + "Y";
					}else{
						upfrm.isb2b.value = upfrm.isb2b.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}
}



function CheckNDoNormal(frm){
	var ret = confirm('저장하시겠습니까?');
	if (ret){
		return true;//frm.submit();
	}
	return false;
}

function CheckNDoitemviewset(frm){
	var ret = confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}


function AnItemlimitsellSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.baedalcd.value = upfrm.baedalcd.value + "|" + frm.baedalcd.value;
					upfrm.limitno.value = upfrm.limitno.value + "|" + frm.limitno.value;
					upfrm.limitsold.value = upfrm.limitsold.value + "|" + frm.limitsold.value;

					if (frm.limityn[0].checked){
						upfrm.limityn.value = upfrm.limityn.value + "|" + "Y";
					}else{
						upfrm.limityn.value = upfrm.limityn.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}

}

function AnBuyPriceSaveFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.sellcash.value = upfrm.sellcash.value + "|" + frm.sellcash.value;
					upfrm.sellvat.value = upfrm.sellvat.value + "|" + frm.sellvat.value;
					upfrm.buycash.value = upfrm.buycash.value + "|" + frm.buycash.value;
					upfrm.buyvat.value = upfrm.buyvat.value + "|" + frm.buyvat.value;
					upfrm.marginrate.value = upfrm.marginrate.value + "|" + frm.marginrate.value;

					if (frm.vtinclude[0].checked){
						upfrm.vtinclude.value = upfrm.vtinclude.value + "|" + "Y";
					}else{
						upfrm.vtinclude.value = upfrm.vtinclude.value + "|" + "N";
					}
				}
			}
		}
		frm.submit();
	}
}

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}


function AnSelectAll(frm,bool){
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
			e.checked = bool;
		}
	}
}

function AnBaljuDetail(frm){
	if (frm.baljumaster.value.length<1){
		alert('발주서를 선택하세요');
		return;
	}
	frm.submit();
}

function AnBaljuSearch(frm){
	frm.baljumaster.value="";
	frm.submit();
}

function AnPopDelivery(){
	window.open('/admin/lib/printdelivery.asp', 'DeliveryPrint', 'width=800,height=600,location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no');
}

function AnDeliverRead(iid,frm){
	frm.id.value = iid;
	frm.action = 'bct_admin_deliver_read.asp';
	frm.submit();
}

function AnOrderView(orderserial){
	alert('준비중...');
}

function AnCheckDelivery(frm){
	var v = confirm('변경하시겠습니까?');
	if (v) {
		frm.mode.value='check';
		frm.submit();
	}
}

function AnDeleteDelivery(frm){
	var v = confirm('삭제하시겠습니까?');
	if (v) {
		frm.mode.value='del';
		frm.submit();
	}
}

function AnWriteDeliveryCom(frm){
	if (frm.tx_com.value.length<1) {
		alert('내용을 입력하세요.');
		frm.tx_com.focus();
		return;
	}

	if (frm.writer.value=='') {
		alert('글쓴이를 선택하세요.');
		return;
	}

	frm.submit();
}

function AnDeliveryWrite(frm){
	if (frm.sitename.value=='') {
		alert('사이트를 선택하세요.');
		return;
	}

	if (frm.buyname.value.length<1) {
		alert('고객명을 입력하세요.');
		frm.buyname.focus();
		return;
	}

	if (frm.orderserial.value.length<1) {
		alert('주문번호를 입력하세요.');
		frm.orderserial.focus();
		return;
	}

	if (frm.writer.value=='') {
		alert('글쓴이를 선택하세요.');
		return;
	}

	if (frm.title.value.length<1) {
		alert('타이틀을 입력하세요.');
		frm.title.focus();
		return;
	}

	if (frm.txmemo.value.length<1) {
		alert('메모를 입력하세요.');
		frm.txmemo.focus();
		return;
	}
	frm.submit();
}

function TnUnderConstruction(){
	alert('UnderConstruction..');
}

function SvNoticeConfirm(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret==true) {
		frm.submit();
	}
}

function SendDreamWinner(frm){
	var pass;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")&&(e.checked)) {
			pass = true;
		}
	}
	if (!pass) {
		alert('당첨자를 선택하세요..');
		return false;
	}
	if (!confirm('당첨자에 추가 하시겠습니까?'))
		return false;
	frm.action = '/admin/lib/dosenddreamwinner.asp';

}

function checkdate3(form){

        var year=form.yyyy2.value;
        var month=form.mm2.value;
        var cal;

                var lastdate=new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

					if((0==year%4&&0!=year%100)||0==year%400){
					  lastdate[1]=29;
					}

 		if(form.dd2.value >= (lastdate[month-1]+1)){
		  alert(month+"월은" + lastdate[month-1] + "일 까지 입니다.\n\n"+lastdate[month-1]+"일 이전날짜로 검색해주세요!");
		}

}

function CheckDateValid(yyyy, mm, dd) {
	var year = yyyy;
    var month = mm;
	var day = dd;

    var lastdate = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

	if(((0 == (year % 4)) && (0 != (year % 100))) || (0 == (year % 400))) {
		lastdate[1] = 29;
	}

 	if (day >= (lastdate[month-1]+1)) {
		alert("에러!!\n\n" + month + " 월은 " + lastdate[month-1] + " 일 까지 입니다.");
		return false;
	}

	return true;
}

function AnItemviewsetSaveAll(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidlist.value = upfrm.itemidlist.value + "|" + frm.itemid.value;
					upfrm.tingpointlist.value = upfrm.tingpointlist.value + "|" + frm.tingpoint.value;
					upfrm.tingpoint_blist.value = upfrm.tingpoint_blist.value + "|" + frm.tingpoint_b.value;
					upfrm.userclasslist.value = upfrm.userclasslist.value + "|" + frm.userclass.value;
					upfrm.limitdivlist.value = upfrm.limitdivlist.value + "|" + frm.limitdiv.value;
					upfrm.limitealist.value = upfrm.limitealist.value + "|" + frm.limitea.value;
					upfrm.selectitemlist.value = upfrm.selectitemlist.value + "|" + frm.selectitem.value;
				}
			}
		}
		frm.submit();
	}
}

function TnFtpUpload2(url,url2){
	window.open('http://partner.10x10.co.kr/ftp/img_upload2.asp?dir=' + url + '&dir2=' + url2, 'ftpupload', 'width=600,height=500,resizable=yes,scrollbars=yes');
}
function TnFtpUpload3(url){
	window.open('http://fiximage.10x10.co.kr/ftp/img_upload.asp?dir=' + url, 'ftpupload', 'width=600,height=500,resizable=yes,scrollbars=yes');
}

function TnFtpUpload(url,url2){
	window.open('http://imgstatic.10x10.co.kr/ftp/img_upload.asp?dir=' + url + '&dir2=' + url2, 'ftpupload_test', 'width=600,height=500,resizable=yes,scrollbars=yes');
}

function GetByteLength(val){
 	var real_byte = val.length;
 	for (var ii=0; ii<val.length; ii++) {
  		var temp = val.substr(ii,1).charCodeAt(0);
  		if (temp > 127) { real_byte++; }
 	}

   return real_byte;
}

//-----------2009 추가 ------------
//입력 키 값이 숫자인지 체크
function onlyNumberInput()
{
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}
function jsTrim(str){
    return str.replace(/^\s\s*/, '').replace(/\s\s*$/, '');
}

// 숫자만 입력받기 IE, FF
function onlyNumber(obj,evt) {
	var evCode = (window.netscape) ? evt.which : event.keyCode ;
	if (isNumeric(evCode)==false)
	{
		alert("숫자만 입력이 가능합니다.");
		if (!window.netscape)
			event.returnValue=false;
		else
			obj.value = obj.value.substr(0,obj.value.length-1);
	}
}

// 숫자와 구분자만 입력받기 IE, FF
function onlyNumberDot(obj,evt) {
	var evCode = (window.netscape) ? evt.which : event.keyCode ;
	if (isNumericDot(evCode)==false)
	{
		alert("숫자만 입력이 가능합니다.");
		if (!window.netscape)
			event.returnValue=false;
		else
			obj.value = obj.value.substr(0,obj.value.length-1);
	}
}

// 숫자인지 체크
function isNumeric( value )
{
	if (value == 189 || value == 109 || value == 8 || value == 9 || value == 13 || (value >= 37 && value <= 40) || value == 46 || value == 27 || value == 116 || (value >= 48 && value <= 57) || (value >= 96 && value <= 105))
		return true;
	else
		return false;
}

// 숫자나 구분자인지 체크
function isNumericDot( value ) {
	if (value == 189 || value == 109 || value == 8 || value == 9 || value == 13 || (value >= 37 && value <= 40) || value == 46 || (value >= 48 && value <= 57) ||  value == 110 || value == 190 || value == 192 || value == 27 || value == 116 || (value >= 96 && value <= 105) || value ==17 || value ==16 || value ==186 || value ==188)
		return true;
	else
		return false;
}

// 폼 필수 필드 유효성 체크
function validField(obj, msg, len)
{
	if (obj.length > 1)
	{
		if (obj[0].type == "radio" || obj[0].type == "checkbox")
		{
			var chk = 0;
			for (var i = 0; i < obj.length; i++)
				if (obj[i].checked)
					chk++;

			if (chk==0)
			{
				if (obj[0].type == "checkbox")
					alert("" + msg + " 하나 이상 체크해주세요.");
				else
					alert("" + msg + " 체크해주세요.");

				obj[0].focus();
				return false;
			}
		}
		else if (obj.type == "select-one")
		{
			if(jsTrim(obj.value) == "")
			{
				alert("" + msg + " 선택해주세요.");
				obj.focus();
				return false;
			}
		}
	}
	else if (obj.type == "radio" || obj.type == "checkbox")
	{
		if (obj.checked==false)
		{
			alert("" + msg + " 체크해주세요.");
			obj.focus();
			return false;
		}
	}
	else
	{
		if(jsTrim(obj.value) == "")
		{
			alert("" + msg + " 입력해주세요.");
			obj.focus();
			return false;
		}
		if (len)
		{
			if (returnByteCount(obj.value) > len)
			{
				alert("" + msg + " 한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
				obj.focus();
				return false;
			}
		}
	}

	return true;
}

// 자리수 형식 체크
function validFieldLeng(obj, msg, len)
{
	if(obj.value.length < len)	// 작으면 false 같거나 크면 상관없음 maxlength로 제한하기 때문에 클 수는 없음
	{
		alert(msg + " 자리수가 맞지 않습니다.\n"+len+" 자리 또는 그 이상으로 입력해주세요.");
		obj.focus();
		return false;
	}
	else
		return true;
}

// 바이트수를 리턴하는 함수
function returnByteCount(val)
{
	var len = val.length;            //현재 value값의 글자 수
	var cnt = 0;                    //한글일 경우 2 그외에는 1바이트 수 저장
	var chr = "";                 //현재 한/영 체크할 letter를 저장

	for (i=0; i<len; i++)
	{
		chr = val.charAt(i);

		// 체크문자가 한글일 경우 2byte 그 외의 경우 1byte 증가
		if (escape(chr).length > 4)
		   cnt += 2;
		else
		   cnt++;
	}
	return cnt;
}

// 폼 필드값 리턴, 조건에 따라 개수가 달라질때
function getFieldValue(obj)
{
	var ret = "";
	if (obj.length > 1)
	{
		if (obj[0].type == "radio" || obj[0].type == "checkbox")
		{
			for (var i = 0; i < obj.length; i++)
				if (obj[i].checked)
					if (ret=="")
						ret = obj[i].value;
					else
						ret += "," + obj[i].value;
		}
		else if (obj.type == "select-one")
		{
			ret = obj.value;
		}
	}
	else
	{
		ret = obj.value;
	}

	return ret;
}

// 팝업창 자동리사이즈, Width를 지정하면 지정한대로
function popupResize(innerWidth)
{
	var strAgent = navigator.userAgent.toLowerCase();
	var strVersion = strAgent.substr(strAgent.indexOf("msie")+5,1);
    var IE	= strAgent.indexOf("MSIE") ?	true : false;
	if (IE)
	{
		var addHeight = (strAgent >=  7) ? 70 : 55;	// 7 이상은 URL창크기만큼 추가

		var innerBody = document.body;
		var innerHeight = innerBody.scrollHeight + (innerBody.offsetHeight - innerBody.clientHeight);
		if (!innerWidth)
			var innerWidth = innerBody.scrollWidth + (innerBody.offsetWidth - innerBody.clientWidth);

		innerWidth += 10;
		innerHeight += addHeight;
		window.resizeTo(innerWidth,innerHeight);
	}
	else					// FF
	{
		var Dwidth = parseInt(document.body.scrollWidth);
		var Dheight = parseInt(document.body.scrollHeight);
		var divEl = document.createElement("div");
		divEl.style.position = "absolute";
		divEl.style.left = "0px";
		divEl.style.top = "0px";
		divEl.style.width = "100%";
		divEl.style.height = "100%";
	    document.body.appendChild(divEl);
	    window.resizeBy(Dwidth-divEl.offsetWidth, Dheight-divEl.offsetHeight);
		document.body.removeChild(divEl);
	}
}


// iframe 길이 자동
function resizeIfr(obj, minHeight) {
	minHeight = minHeight || 10;

	try {
		var getHeightByElement = function(body) {
			var last = body.lastChild;
			try {
				while (last && last.nodeType != 1 || !last.offsetTop) last = last.previousSibling;
				return last.offsetTop+last.offsetHeight;
			} catch(e) {
				return 0;
			}

		}

		var doc = obj.contentDocument || obj.contentWindow.document;
		if (doc.location.href == 'about:blank') {
			obj.style.height = minHeight+'px';
			return;
		}

		//var h = Math.max(doc.body.scrollHeight,getHeightByElement(doc.body));
		//var h = doc.body.scrollHeight;
		if (/MSIE/.test(navigator.userAgent)) {
			var h = doc.body.scrollHeight;
		} else {
			var s = doc.body.appendChild(document.createElement('DIV'))
			s.style.clear = 'both';

			var h = s.offsetTop;
			s.parentNode.removeChild(s);
		}

		//if (/MSIE/.test(navigator.userAgent)) h += doc.body.offsetHeight - doc.body.clientHeight;
		if (h < minHeight) h = minHeight;

		obj.style.height = h + 'px';
		if (typeof resizeIfr.check == 'undefined') resizeIfr.check = 0;
		if (typeof obj._check == 'undefined') obj._check = 0;

//		if (obj._check < 5) {
//			obj._check++;
			setTimeout(function(){ resizeIfr(obj,minHeight) }, 200); // check 5 times for IE bug
//		} else {
			//obj._check = 0;
//		}
	} catch (e) {
		//alert(e);
	}

}


//동영상 SerialNumber 검색 팝업창
function jsSearchVideoSn(frmName,compName,vDiv){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/sitemaster/popVideoSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal+ "&vDiv=" + vDiv,"popVideoSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}



function plusComma(num){
	if (num < 0) { num *= -1; var minus = true}
	else var minus = false

	var dotPos = (num+"").split(".")
	var dotU = dotPos[0]
	var dotD = dotPos[1]
	var commaFlag = dotU.length%3

	if(commaFlag) {
		var out = dotU.substring(0, commaFlag)
		if (dotU.length > 3) out += ","
	}
	else var out = ""

	for (var i=commaFlag; i < dotU.length; i+=3) {
		out += dotU.substring(i, i+3)
		if( i < dotU.length-3) out += ","
	}

	if(minus) out = "-" + out
	if(dotD) return out + "." + dotD
	else return out
}

//바코드 OCX
function drawTTPprintOcx(iname, iversion){
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "      name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:E0CFD990-7055-4BFD-927E-BF5553AF0F54'";
    iObjStr = iObjStr + "	  codebase='http://scm.10x10.co.kr/common/cab/tenTTPBarPrint.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

function drawTTPprintOcxV2(iname, iversion){
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "      name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:4B4DE9A2-A9B5-403B-8AFF-4967823E3BB2'";
    iObjStr = iObjStr + "	  codebase='http://scm.10x10.co.kr/common/cab/TenTTPBar.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

function drawSrp350PlotOcx(iname, iversion){
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "     id='" + iname + "' name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:5DC34DA8-9C0F-4A43-B772-78090A204600'";
    iObjStr = iObjStr + "	  codebase='http://scm.10x10.co.kr/common/cab/TnSRPPlot.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

//우편번호 입력
function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;

    eval(frmname + ".zipaddr").value = addr;
    eval(frmname + ".useraddr").value = dong;
}

// 날짜크기 비교
function dateChkComp(dt1,dt2)
{
	//구분자로 나누어 배열로 변환
	v0=dt1.split("-");
	v1=dt2.split("-");

	//일자에 해당하는 타임스탬프로 변환
	v0=new Date(v0[0],v0[1],v0[2]).valueOf();
	v1=new Date(v1[0],v1[1],v1[2]).valueOf();

	//일차이를 구한뒤 하루에 해당하는 값으로 곱하여, 초단위를 일단위로 변환
	cha=(v1-v0)/(1000*60*60*24);

	if(cha>0)
		return true;
	else
		return false;
}

//매장직원 매장 지정
function shopreg(empno){
	var shopreg = window.open('/common/offshop/member/shopuser_reg.asp?empno='+empno,'shopreg','width=800,height=400,scrollbars=yes,resizable=yes');
	shopreg.focus();
}

//asp instr 함수 생성	//2011.02.11 한용민 생성
function instr(strsearch,charsearchfor){
	for ( var i =0; i < strsearch.length; i++){
		if (charsearchfor == mid(strsearch , i ,1) ){
			return i;
		}
	}
	return -1;
}

//asp mid 함수생성	//2011.02.11 한용민 생성
function mid(str ,start,len){
	if (start < 0 || len < 0) return '';
	var iend;
	var ilen = str.length ;

	if( start + len > ilen){
		iend = ilen;
	} else {
		iend = start + len;
	}

	return str.substring(start,iend);
}

//문자열의 공백여부 체크 2011.03.16 정윤정 생성
function jsChkBlank(str)
{
    if (str == "" || str.split(" ").join("") == ""){
        return true;
	}
    else{
        return false;
	}
}

//TTP-243 바코드 프린트	팝업		'//2013.03.04 한용민 생성
//listgubun구분 : 상품리스트=ITEM , 업체주문리스트=JUMUN
//idx구분 : 업체주문 입출코드
//prdcode : 물류코드
function PopBarCodettpPrint(listgubun, idx, prdcode, makerid, shopid){
	var PopBarCodettpPrint = window.open('/common/popBarcodePrintOffline.asp?listgubun='+listgubun+'&ipchul='+idx+'&prdcode='+prdcode+'&makerid='+makerid+'&shopid='+shopid,'PopBarCodettpPrint','width=1024,height=768,scrollbars=yes,resizable=yes');
	PopBarCodettpPrint.focus();
}


//상품 미리보기
function jsGoPreItem(wwwURL, itemid){
	 window.open('about:blank').location.href = wwwURL+"/shopping/category_prd.asp?itemid="+itemid; 
}