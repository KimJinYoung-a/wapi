
// =============================================================================
// �Ʒ� �ΰ��� ������ ��� ������ �־�� �Ѵ�.
// ������	/js/ttpbarcode.js
// SCM		/js/ttpbarcode.js
// =============================================================================

/*

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>

// �ѻ�ǰ ���ڵ� ���
// <input type="button" class="button" value="���" onClick="BarcodePrint('102140800012', '122kcal', 'roll (pencil case)', 'carrot(orange)', '10,000', 5)">
function BarcodePrint(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	// /js/barcode.js ����
	if (initTTPprinter("TTP-243_45x22", "T", "Y", "www.10x10.co.kr", "Y", "��", "Y", 3, 0) != true) {
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.');
		return;
	}

	if (printno*1 < 1) {
		alert("������ 0 �Դϴ�.");
		return;
	}

	printTTPOneBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno);
}

// ������ǰ ���ڵ� ���
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
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	// /js/barcode.js ����
	if (initTTPprinter("TTP-243_45x22", "G", "Y", "www.10x10.co.kr", "Y", "��", "Y", 3, 0) != true) {
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[4]');
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
var TTP_BARCODETYPE					// T or G(�ٹ����� ���ڵ� or ������ڵ�)
var TTP_SHOWDOMAINYN				// y or n
var TTP_DOMAINNAME					// www.10x10.co.kr
var TTP_SHOWPRICEYN					// y or n
var TTP_CURRENCYCHAR				// ��(\ �������� �ƴ�) or $ or ��
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
		alert('�������� �ʴ� �����Դϴ�.(TTP-243_45x22, TTP-243_80x50, TTP-index243_80x50 �� ����)');
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
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[1]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// ���� Ŭ����..  ���Ұ��.. ù��° ������ ������ ��� ���Ƽ�..���ļ� ����;;
	iTTPBar.AXclearbuffer();

	if ((skipnotinserted != true) && (itemname == "")) {
		if (confirm("��ǰ���� �Էµ��� ���� ��ǰ�� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?") != true) {
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
	    	//�ɼ��ڵ忡 Z�� ��� �������
	    	iTTPBar.AXbarcode('30', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else if (TTP_BARCODETYPE == "T") {
	    	// �ٹ����� �Ϲ� �����ڵ�
	    	iTTPBar.AXbarcode('50', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
	    } else {
	    	// ������ڵ�
	    	iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
			iTTPBar.AXwindowsfont(50,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', barcode );
	    }

	    // printno �� ����Ʈ
	    iTTPBar.AXprintlabel('1', printno*1);

	    iTTPBar.AXformfeed();
	}

	iTTPBar.AXcloseport();
}

function printTTPOneIndexBarcode(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	var skipnotinserted = false;

	if (TTP_INITIALIZED != true) {
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[5]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	// ���� Ŭ����..  ���Ұ��.. ù��° ������ ������ ��� ���Ƽ�..���ļ� ����;;
	iTTPBar.AXclearbuffer();

	if ((skipnotinserted != true) && (itemname == "")) {
		if (confirm("��ǰ���� �Էµ��� ���� ��ǰ�� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?") != true) {
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
	        iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '�Һ��ڰ� : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
	    }

	    iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

	    if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
	    	//�ɼ��ڵ忡 Z�� ��� �������
	    	iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else if (TTP_BARCODETYPE == "T") {
	    	// �ٹ����� �Ϲ� �����ڵ�
	    	iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
	    } else {
	    	// ������ڵ�
	    	// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
	    }

	    iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, 'EAN128', '30', '0', '0', '2', '4',barcode);

	    // printno �� ����Ʈ
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
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// ���� Ŭ����..  ���Ұ��.. ù��° ������ ������ ��� ���Ƽ�..���ļ� ����;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		if ((skipnotinserted != true) && (itemname == "")) {
			if (confirm("��ǰ���� �Էµ��� ���� ��ǰ�� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?") != true) {
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
		    	//�ɼ��ڵ忡 Z�� ��� �������
		    	iTTPBar.AXbarcode('30', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
		    } else if (TTP_BARCODETYPE == "T") {
		    	// �ٹ����� �Ϲ� �����ڵ�
		    	iTTPBar.AXbarcode('50', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', getTTPBarcodeString(barcode) );
		    } else {
		    	// ������ڵ�
		    	iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
		    	iTTPBar.AXwindowsfont(80,140 + TTP_HEIGHTOFFSET,25,0,0,0,'Arial', barcode);  //20121219 �߰�
		    }

		    // printno �� ����Ʈ
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
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[2]');
		return;
	}

    iTTPBar.AXclearbuffer();
    iTTPBar.AXsendcommand('DIRECTION 1');
    iTTPBar.AXsetup(TTP_PAPERWIDTH, TTP_PAPERHEIGHT, '2', '10', '0', TTP_PAPERMARGIN, '0');

	for (var i = 0; i < arrObject.length; i++) {
		// ���� Ŭ����..  ���Ұ��.. ù��° ������ ������ ��� ���Ƽ�..���ļ� ����;;
		iTTPBar.AXclearbuffer();

		v = arrObject[i];

		barcode = v.barcode;
		makerid = v.makerid;
		itemname = v.itemname;
		itemoptionname = v.itemoptionname;
		customerprice = v.customerprice;
		printno = v.printno;

		if ((skipnotinserted != true) && (itemname == "")) {
			if (confirm("��ǰ���� �Էµ��� ���� ��ǰ�� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?") != true) {
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
				iTTPBar.AXwindowsfont(50, 170 + TTP_HEIGHTOFFSET, 30, 0, 0, 0, 'Arial', '�Һ��ڰ� : ' + TTP_CURRENCYCHAR + ' ' + customerprice);
			}

			iTTPBar.AXwindowsfont(180, 220 + TTP_HEIGHTOFFSET, 110, 0, 2, 0, 'Arial', getTTPBarcodeItemidString(barcode) );

			if ((TTP_BARCODETYPE == "T") && (InStr(barcode, 'Z') >= 0)) {
				//�ɼ��ڵ忡 Z�� ��� �������
				iTTPBar.AXwindowsfont(370,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else if (TTP_BARCODETYPE == "T") {
				// �ٹ����� �Ϲ� �����ڵ�
				iTTPBar.AXwindowsfont(340,330 + TTP_HEIGHTOFFSET,40,0,0,0,'Arial', barcode);
			} else {
				// ������ڵ�
				// iTTPBar.AXbarcode('20', 100 + TTP_HEIGHTOFFSET,'EAN128','30','0','0','2','4',barcode);
			}

			iTTPBar.AXbarcode('50', 335 + TTP_HEIGHTOFFSET, 'EAN128', '30', '0', '0', '2', '4',barcode);

			// printno �� ����Ʈ
			iTTPBar.AXprintlabel('1', printno*1);

			// ����� �ʰ� ����Ѵ�.
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
// 100�� �̻�
// 10010000000000 => 01000000
// 1000000 => 01000000
// =============================================================================
// 100�� �̸�
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
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[5]');
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

// TTP ��� ��ġ
drawTTPprintOcxV2__('iTTPBar','1,0,0,3');
