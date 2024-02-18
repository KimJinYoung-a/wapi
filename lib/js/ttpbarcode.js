
// =============================================================================
// 아래 두개의 파일을 모두 수정해 주어야 한다.
// 로직스	/js/ttpbarcode.js
// SCM		/js/ttpbarcode.js
// =============================================================================

/*

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>

// 한상품 바코드 출력
// <input type="button" class="button" value="출력" onClick="BarcodePrint('102140800012', '122kcal', 'roll (pencil case)', 'carrot(orange)', '10,000', 5)">
function BarcodePrint(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_45x22", "T", "Y", "www.10x10.co.kr", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.');
		return;
	}

	if (printno*1 < 1) {
		alert("수량이 0 입니다.");
		return;
	}

	printTTPOneBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno);
}

// 여러상품 바코드 출력
function BarcodePrintSelected() {
	var frmdetail = document.frmdetail;
	var arr = new Array();
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;

	for (var i = 0; i < frmdetail.chk.length; i++) {
		if (frmdetail.chk[i].type == "checkbox") {
			if (frmdetail.chk[i].checked) {
				barcode			= frmdetail.itembarcode[i].value;

				makerid			= frmdetail.makerid[i].value;
				itemname		= frmdetail.itemname[i].value;
				itemoptionname	= frmdetail.itemoptionname[i].value;
				customerprice	= frmdetail.customerprice[i].value;
				printno			= frmdetail.checkitemno[i].value;

				var v = new TTPBarcodeDataClass(barcode, makerid, itemname, itemoptionname, customerprice, printno);
				arr.push(v);
			}
		}
	}

	if (arr.length < 1) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_45x22", "G", "Y", "www.10x10.co.kr", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]');
		return;
	}

	printTTPMultiBarcode(arr);
}

</script>

*/

function checkTTPprinterExist() {
    if (!iTTPBar.AXopenport(TTP_PRINTERTYPE)) {
        alert("a" + TTP_PRINTERTYPE);
        return false;
    }

    return true;
}

function InStr(str, substr, start) {
	var oStr = new String(str);
	return oStr.indexOf(substr,start);
}

function TTPBarcodeDataClass(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
    this.barcode = barcode;
    this.makerid = makerid;
    this.itemname = itemname;
    this.itemoptionname = itemoptionname;
    this.customerprice = customerprice;
    this.printno = printno;
}

// ============================================================================
var TTP_INITIALIZED = false			// true or false
var TTP_TTPTYPE						// TTP-243_45x22
var TTP_PRINTERTYPE					// TTP-243
var TTP_BARCODETYPE					// T or G(텐바이텐 바코드 or 범용바코드)
var TTP_SHOWDOMAINYN				// y or n
var TTP_DOMAINNAME					// www.10x10.co.kr
var TTP_SHOWPRICEYN					// y or n
var TTP_CURRENCYCHAR				// ￦(\ 역슬래시 아님) or $ or ￥
var TTP_SHOPBRANDYN					// y or n
var TTP_PAPERWIDTH					// 45
var TTP_PAPERHEIGHT					// 22
var TTP_PAPERMARGIN					// 3
var TTP_HEIGHTOFFSET				// 0
// ============================================================================

function initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) {
	var s1, s2;

	TTP_INITIALIZED = false;

	if ((ttptype != "TTP-243_45x22") && (ttptype != "TTP-243_80x50") && (ttptype != "TTP-index243_80x50")) {
		alert('지원하지 않는 형식입니다.(TTP-243_45x22, TTP-243_80x50, TTP-index243_80x50 만 지원)');
		return false;
	}

	s1 = ttptype.split("_");
	s2 = s1[1].split("x");
	TTP_TTPTYPE			= ttptype;
	TTP_PRINTERTYPE		= s1[0];
	TTP_PAPERWIDTH		= s2[0]*1;
	TTP_PAPERHEIGHT		= s2[1]*1;

	TTP_BARCODETYPE		= barcodetype;
	TTP_SHOWDOMAINYN	= showdomainyn;
	TTP_DOMAINNAME		= domainname;
	TTP_SHOWPRICEYN		= showpriceyn;
	TTP_CURRENCYCHAR	= currencychar;
	TTP_SHOPBRANDYN		= shopbrandyn;
	TTP_PAPERMARGIN		= papermargin;
	TTP_HEIGHTOFFSET	= heightoffset;

	if (checkTTPprinterExist() != true) {
        return false;
	}

	TTP_INITIALIZED 	= true;

	return true;
}

function printTTPOneBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[1]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	if ((skipnotinserted != true) && (itemname == "")) {
		if (confirm("상품명이 입력되지 않은 상품이 있습니다. 진행하시겠습니까?") != true) {
			return;
		}
	}
	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(75, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', TTP_DOMAINNAME);
		}

		if (itemoptionname == "") {
			iTTPBar.AXwindowsfont(20, 45 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname);
		} else {
			iTTPBar.AXwindowsfont(20, 45 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
		}

	    if (TTP_SHOWPRICEYN == "Y"){
	        iTTPBar.AXwindowsfont(260, 65 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
	    }

	    if (TTP_SHOPBRANDYN == "Y"){
			iTTPBar.AXwindowsfont(20, 65 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', makerid);
	    }

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXbarcode('30', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXbarcode('50', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else {
	    	// 범용바코드
	    	iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
			iTTPBar.AXwindowsfont(50,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', barcode );
	    }

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

function printTTPOneIndexBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
	iTTPBar.AXclearbuffer();

	if ((skipnotinserted != true) && (itemname == "")) {
		if (confirm("상품명이 입력되지 않은 상품이 있습니다. 진행하시겠습니까?") != true) {
			return;
		}
	}
	skipnotinserted = true;

	if ((itemname != "") && printno*1 > 0) {
		if (TTP_SHOWDOMAINYN == "Y") {
			iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', TTP_DOMAINNAME);
		}

	    if (TTP_SHOPBRANDYN == "Y"){
			iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', makerid);
	    }

		if (itemoptionname == "") {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
		} else {
			iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
		}

	    if (TTP_SHOWPRICEYN == "Y"){
	        iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '소비자가 : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
	    }

	    iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//옵션코드에 Z가 들어 있을경우
	    	iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else if (TTP_BARCODETYPE == "T") {
	    	// 텐바이텐 일반 물류코드
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else {
	    	// 범용바코드
	    	// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    }

	    iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, 'EAN128', '30', '0', '0', '2', '4',barcode);

	    // printno 장 프린트
	    iTTPBar.AXprintlabel('1', printno*1);

	    iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

function printTTPMultiBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		if ((skipnotinserted != true) && (itemname == "")) {
			if (confirm("상품명이 입력되지 않은 상품이 있습니다. 진행하시겠습니까?") != true) {
				return;
			}
		}
		skipnotinserted = true;
		if ((itemname != "") && (printno*1 > 0)) {
			if (TTP_SHOWDOMAINYN == "Y") {
				iTTPBar.AXwindowsfont(75, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', TTP_DOMAINNAME);
			}

			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(20, 45 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname);
			} else {
				iTTPBar.AXwindowsfont(20, 45 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
			}

		    if (TTP_SHOWPRICEYN == "Y"){
		        iTTPBar.AXwindowsfont(260, 65 + TTP_HEIGHTOFFSET, 20, 0, 2, 0, 'Arial', TTP_CURRENCYCHAR + ' ' + customerprice);
		    }

		    if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(20, 65 + TTP_HEIGHTOFFSET, 20, 0, 0, 0, 'Arial', makerid);
		    }

		    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
		    	//옵션코드에 Z가 들어 있을경우
		    	iTTPBar.AXbarcode('30', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
		    } else if (TTP_BARCODETYPE == "T") {
		    	// 텐바이텐 일반 물류코드
		    	iTTPBar.AXbarcode('50', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
		    } else {
		    	// 범용바코드
		    	iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', barcode);  //20121219 추가
		    }

		    // printno 장 프린트
		    iTTPBar.AXprintlabel('1', printno*1);

		    iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

function printTTPMultiIndexBarcode(arrObject) {
	var skipnotinserted = false;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var v;

	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// 버퍼 클리어..  안할경우.. 첫번째 찍히는 내역이 계속 남아서..겹쳐서 찍힘;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		if ((skipnotinserted != true) && (itemname == "")) {
			if (confirm("상품명이 입력되지 않은 상품이 있습니다. 진행하시겠습니까?") != true) {
				return;
			}
		}
		skipnotinserted = true;
		if ((itemname != "") && (printno*1 > 0)) {
			if (TTP_SHOWDOMAINYN == "Y") {
				iTTPBar.AXwindowsfont(50, 0 + TTP_HEIGHTOFFSET, 30, 0, 2, 1, 'Arial', TTP_DOMAINNAME);
			}

			if (TTP_SHOPBRANDYN == "Y"){
				iTTPBar.AXwindowsfont(50, 40 + TTP_HEIGHTOFFSET, 30, 0, 2, 0, 'Arial', makerid);
			}

			if (itemoptionname == "") {
				iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname);
			} else {
				iTTPBar.AXwindowsfont(50, 130 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', itemname + " - " + itemoptionname);
			}

			if (TTP_SHOWPRICEYN == "Y"){
				iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '소비자가 : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
			}

			iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

			if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
				//옵션코드에 Z가 들어 있을경우
				iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else if (TTP_BARCODETYPE == "T") {
				// 텐바이텐 일반 물류코드
				iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else {
				// 범용바코드
				// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
			}

			iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, 'EAN128', '30', '0', '0', '2', '4',barcode);

			// printno 장 프린트
			iTTPBar.AXprintlabel('1', printno*1);

			// 띄우지 않고 출력한다.
			// iTTPBar.AXformfeed();
		}

	}

	iTTPBar.AXcloseport();
}

// 10010000000000 => 10-01000000-0000
// 104444440000 => 10-444444-0000
function getTTPBarcodeString(barcode) {
	var itemgubun, itemid, itemoption;

	itemgubun 	= barcode.substring(0, 2);
	itemid 		= barcode.substring(2, (barcode.length - 4));
	itemoption	= barcode.substring((barcode.length - 4), barcode.length);

	return itemgubun + "-" + itemid + "-" + itemoption;
}

// =============================================================================
// 100만 이상
// 10010000000000 => 01000000
// 1000000 => 01000000
// =============================================================================
// 100만 미만
// 100444440000 => 044444
// 444444 => 0444444
// =============================================================================
function getTTPBarcodeItemidString(barcode) {
	var itemgubun, itemid, itemoption;

	if (barcode.length >= 12) {
		itemid 		= barcode.substring(2, (barcode.length - 4));
	} else {
		if ((barcode*1) >= 1000000) {
			itemid = (100000000 + barcode*1) + "";
			itemid = itemid.substring((itemid.length - 8), itemid.length);
		} else {
			itemid = (1000000 + barcode*1) + "";
			itemid = itemid.substring((itemid.length - 6), itemid.length);
		}
	}

	return itemid;
}

function printTTPInnerBoxBarcode(baljudate, baljuid, boxno, innerboxweight, innerboxbarcode, innerboxbarcodeforshow) {
	if (TTP_INITIALIZED != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand("DIRECTION 1");
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, "2", "10", "0", TTP_PAPERMARGIN, "0");

	iTTPBar.AXwindowsfont(30,0 + TTP_HEIGHTOFFSET,40,0,2,1,"Arial","                INNER BOX INDEX               ");
	iTTPBar.AXwindowsfont(30,55 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","SHOPID : " + baljuid);
	iTTPBar.AXwindowsfont(30,90 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","                " + TTP_DOMAINNAME);
	iTTPBar.AXwindowsfont(30,130 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","DATE : " + baljudate);
	iTTPBar.AXwindowsfont(30,170 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","INNER BOX NO. : " + boxno);
	iTTPBar.AXwindowsfont(30,210 + TTP_HEIGHTOFFSET,40,0,0,0,"Arial","INNER BOX WEIGHT : " + innerboxweight + " KG");
	iTTPBar.AXbarcode("160",280 + TTP_HEIGHTOFFSET,"EAN128","40","0","0","2","4", innerboxbarcode);
	iTTPBar.AXwindowsfont(30,345 + TTP_HEIGHTOFFSET,30,0,0,0,"Arial","                      " + innerboxbarcodeforshow);

	iTTPBar.AXprintlabel("1","1");
	iTTPBar.AXcloseport();
}

function drawTTPprintOcxV2__(iname, iversion) {
    var iObjStr = "";
    iObjStr = "<OBJECT"
    iObjStr = iObjStr + "      name='" + iname + "'";
    iObjStr = iObjStr + "	  classid='clsid:4B4DE9A2-A9B5-403B-8AFF-4967823E3BB2'";
    iObjStr = iObjStr + "	  codebase='http://logics.10x10.co.kr/common/cab/TenTTPBar.cab#version=" + iversion + "'";
    iObjStr = iObjStr + "	  width=0";
    iObjStr = iObjStr + "	  height=0";
    iObjStr = iObjStr + "	  align=center";
    iObjStr = iObjStr + "	  hspace=0";
    iObjStr = iObjStr + "	  vspace=0";
    iObjStr = iObjStr + ">";
    iObjStr = iObjStr + "</OBJECT>";

    document.write(iObjStr);
}

// TTP 모듈 설치
drawTTPprintOcxV2__('iTTPBar','1,0,0,3');
