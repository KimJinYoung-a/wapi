<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/adminbodyhead.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/cjmall/cjmallitemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, cjmallitemid, extsellyn, ExtNotReg, isReged, MatchCate, MatchPrddiv
Dim expensive10x10, diffPrc, cjmallYes10x10No, cjmallNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oCJMall
page    				= request("page")
research				= request("research")
itemid  				= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")
infoDiv					= request("infoDiv")
extsellyn				= request("extsellyn")
cjmallitemid			= request("cjmallitemid")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
cjmallYes10x10No		= request("cjmallYes10x10No")
cjmallNo10x10Yes		= request("cjmallNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp) 
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'CJMall ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If cjmallitemid<>"" then
	Dim iA2, arrTemp2, arrcjmallitemid
	cjmallitemid = replace(cjmallitemid,",",chr(10))
	cjmallitemid = replace(cjmallitemid,chr(13),"")
	arrTemp2 = Split(cjmallitemid,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2) 
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrcjmallitemid = arrcjmallitemid & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	cjmallitemid = left(arrcjmallitemid,len(arrcjmallitemid)-1)
End If

SET oCJMall = new CCjmall
	oCJMall.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oCJMall.FPageSize					= 50
Else
	oCJMall.FPageSize					= 20
End If
	oCJMall.FRectCDL					= request("cdl")
	oCJMall.FRectCDM					= request("cdm")
	oCJMall.FRectCDS					= request("cds")
	oCJMall.FRectItemID					= itemid
	oCJMall.FRectItemName				= itemname
	oCJMall.FRectSellYn					= sellyn
	oCJMall.FRectLimitYn				= limityn
	oCJMall.FRectSailYn					= sailyn
	oCJMall.FRectonlyValidMargin		= onlyValidMargin
	oCJMall.FRectMakerid				= makerid
	oCJMall.FRectCJMallPrdNo			= cjmallitemid
	oCJMall.FRectMatchCate				= MatchCate
	oCJMall.FRectPrdDivMatch			= MatchPrddiv
	oCJMall.FRectIsMadeHand				= isMadeHand
	oCJMall.FRectIsOption				= isOption
	oCJMall.FRectIsReged				= isReged

	oCJMall.FRectExtNotReg				= ExtNotReg
	oCJMall.FRectExpensive10x10			= expensive10x10
	oCJMall.FRectdiffPrc				= diffPrc
	oCJMall.FRectCjmallYes10x10No		= cjmallYes10x10No
	oCJMall.FRectCjmallNo10x10Yes		= cjmallNo10x10Yes
	oCJMall.FRectExtSellYn				= extsellyn
	oCJMall.FRectInfoDiv				= infoDiv
	oCJMall.FRectFailCntOverExcept		= ""
	oCJMall.FRectFailCntExists			= failCntExists
	oCJMall.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oCJMall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oCJMall.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oCJMall.getCjmallreqExpireItemList
Else
	oCJMall.getCjmallRegedItemList			'�� �� ����Ʈ
End If
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}

function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "J"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "J"
   			comp.form.ExtNotReg.disabled = true;
			comp.form.extsellyn.value = "N";
			comp.form.sellyn.value = "Y";
   		}else{
			if (document.getElementById("NR").checked == false){
				comp.form.extsellyn.value = "Y";
			}else{
				comp.form.extsellyn.value = "";
				comp.form.sellyn.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="cjmallYes10x10No")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.isReged.checked = true;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="cjmallNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
    	}
    }
    
    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.cjmallYes10x10No.checked){
            comp.form.cjmallYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="";
	        comp.form.extsellyn.value = "Y";
    	}
    }
	if ((comp.name=="diffPrc")){
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
        }
	}

	if (comp.name=="reqEdit"){
		if (comp.checked){
			document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="cjmallYes10x10No")&&(frm.cjmallYes10x10No.checked)){ frm.cjmallYes10x10No.checked=false }
	if ((comp.name!="cjmallNo10x10Yes")&&(frm.cjmallNo10x10Yes.checked)){ frm.cjmallNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}

//��Ͽ��� ���� Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
}
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("<%=manageUrl%>/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=cjmall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('<%=manageUrl%>/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=cjmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/outmall/cjmall/popcjmallCateList.asp","popCateMancjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//��ǰ�з� ����
function pop_prdDivManager() {
	var pCM2 = window.open("/outmall/cjmall/popcjmallprdDivList.asp","popprdDivcjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//��ǰ ���
function CjSelectRegProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
	if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "RegSelect";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//��ǰ ��ȸ
function checkCjItemConfirm(comp) {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
	document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItem";
	document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}
//��ǰ ���� ����
function CjmallSellYnProcess(chkYn) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="�Ǹ���";break;
		case "N": strSell="�Ͻ��ߴ�";break;
		case "X": strSell="�Ǹ�����(����)";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��cjmall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� cjmall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//���� ����
function CjSelectEditProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n�ؿɼ��߰� �� ��ǰ ������ �����˴ϴ�. ���� �� ���´� �������� �ʽ��ϴ�')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ �ǸŻ��� ����
function CjSelectSaleStatEditProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm('CjMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ���¸� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CjMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditDanpum").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdSaleDTSel";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}
//��ǰ ���� ����
function CjSelectQTYEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditqty").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditQty";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ ���� ����
function CjSelectPriceEditProcess2() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect2";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ ���� ����
function CjSelectPriceEditProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
       // document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//�����ڵ� �˻�
function popCjCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	xLink.location.href = "/outmall/cjmall/actCjMallReq.asp?cmdparam=cjmallCommonCode&CommCD="+ccd+"";
}
//����+��ǰ ����
function CjSelectEdit2Process() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		//document.getElementById("btnEditSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSelect2";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//���û�ǰ ����Ȯ�� �� �ǸŻ��� check - batch
function batchStatCheck(){
    document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItemAuto";
	document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/outmall/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=cjmall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

function CjSellynSubmit(yn){

	if(yn == true){
	    if (document.getElementById('theday').value.length!=10) {
	        alert('��¥ �������� �Է��� �ּ���.yyyy-mm-dd');
	        return;
	     }
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST&sday="+document.getElementById('theday').value+"";
	}else{
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday="+document.getElementById('sday').value+"";
	}
	document.getElementById("btnSell1").disabled=true;
	document.getElementById("btnSell2").disabled=true;
}
//���� ��ũ�� scm���� ������ ���� : �̳���� �������� ���̼����� wapi.10x10.co.kr���� ��� �۵����ϱ� ����
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('http://scm.10x10.co.kr/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;<% OutmallAdminInfo("cjmall") %>
		&nbsp;
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		cjmall ��ǰ�ڵ� : <textarea rows="2" cols="20" name="cjmallitemid" id="itemid"><%=replace(cjmallitemid,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/lib/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� : 
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >CJmall ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >CJmall ��Ͽ����̻�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >CJmall ��Ͽ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >CJmall ���۽õ��߿���
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >CJmall ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >CJmall ��ϿϷ�(����)
		</select>&nbsp;
		<label><input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">��ü</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">�̵��</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">ǰ��ó�����</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">��ϻ�ǰ �ǸŰ���</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">��Ͽ�������Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/outmall/incsearch1.asp"-->
		ī�װ�
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;
		��ǰ�з�
		<select name="MatchPrddiv" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>cjmall ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="cjmallYes10x10No" <%= ChkIIF(cjmallYes10x10No="on","checked","") %> ><font color=red>cjmall�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="cjmallNo10x10Yes" <%= ChkIIF(cjmallNo10x10Yes="on","checked","") %> ><font color=red>cjmallǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<form name="frmReg" method="post" action="cjmallItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
			</td>
			<td align="right">
				<font color="RED">���� 2�� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="cjMall��ǰ�з���Ī" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="cjMallī�װ���Ī" onclick="pop_CateManager();">&nbsp;&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
<!--	<center><strong>�ɼ��߰��ݾ� �ִ� ��ǰ�� ���ݼ��� �� :</strong> <font color="red">��ǰ ���� ����</font> �Ϸ� ��  <font color="red">��ǰ ����(����)</font> ����</center> -->
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
	    		������ǰ ��� :
				<input class="button" type="button" id="btnRegSel" value="��ǰ ���" onClick="CjSelectRegProcess(true);" <%= Chkiif(isReged = "A", "disabled", "") %> >&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="���� ����" onClick="CjSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDanpum" value="��ǰ ����(����)" onClick="CjSelectSaleStatEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditqty" value="��ǰ ����(����)" onClick="CjSelectQTYEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="��ǰ ����(����)" onClick="CjSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDate" value="����+��ǰ ����" onClick="CjSelectEdit2Process();">
				<!--
				<input class="button" type="button" id="btnEditDate" value="���û�ǰ �������" onClick="CjSelectDateEditProcess();">
				-->
				<br><br>
				���ο��� �˻� :
				<!--
				<input type="text" name="theday" value="" size="10" maxlength="10">
				<input class="button" type="button" id="btnSell1" value="Ư����¥ ���ο��� Ȯ��" onClick="CjSellynSubmit(true);">&nbsp;&nbsp;
				<select name="sday" class="select" id="sday">
				<% For i = 0 to 9 %>
					<option value="<%=i%>"><%=i%>
				<% Next %>
				</select>����&nbsp;
				<input class="button" type="button" id="btnSell2" value="���� �Ⱓ ���ο��� Ȯ��" onClick="CjSellynSubmit(false);" >
				-->
				<input class="button" type="button" id="btnSelectDate" value="��ǰ ��ȸ" onClick="checkCjItemConfirm(this);" >
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="L126">�ù���ڵ�
					<option value="6009">����Ÿ��
					<option value="8047">�����ä�α���
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="popCjCommCDSubmit();" >
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ͻ��ߴ�</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="CjmallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input class="button" type="button" value="�ǸŻ���Check(������)" onClick="batchStatCheck();">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- �׼� �� -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oCJMall.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCJMall.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">CJMall�����<br>CJMall����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">CJMall<br>���ݹ��Ǹ�</td>
	<td width="70">CJMall<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">��ǰ�з�<br>��Ī����</td>
</tr>
<% For i = 0 To oCJMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oCJMall.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oCJMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oCJMall.FItemList(i).FItemID %>','cjMall','<%=oCJMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center"><a href="<%=wwwURL%>/<%=oCJMall.FItemList(i).FItemID%>" target="_blank"><%= oCJMall.FItemList(i).FItemID %></a><br><%= oCJMall.FItemList(i).getcjmallStatName %></td>
	<td align="left"><%= oCJMall.FItemList(i).FMakerid %><%= oCJMall.FItemList(i).getDeliverytypeName %><br><%= oCJMall.FItemList(i).FItemName %></td>
	<td align="center"><%= oCJMall.FItemList(i).FRegdate %><br><%= oCJMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oCJMall.FItemList(i).FcjmallRegdate %><br><%= oCJMall.FItemList(i).FcjmallLastUpdate %></td>
	<td align="right">
	<% If oCJMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oCJMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oCJMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oCJMall.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oCJMall.FItemList(i).Fbuycash/oCJMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).IsSoldOut Then
			If oCJMall.FItemList(i).FSellyn = "N" Then
	%>
			<font color="red">ǰ��</font>
	<%
			Else
	%>
			<font color="red">�Ͻ�<br>ǰ��</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).FItemdiv = "06" OR oCJMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oCJMall.FItemList(i).FcjmallStatCd > 0) Then
			If Not IsNULL(oCJMall.FItemList(i).FcjmallPrice) Then
				If (oCJMall.FItemList(i).Fsellcash <> oCJMall.FItemList(i).FcjmallPrice) Then
	%>
					<strong><%= formatNumber(oCJMall.FItemList(i).FcjmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oCJMall.FItemList(i).FcjmallPrice,0)&"<br>"
				End If

				If (oCJMall.FItemList(i).FSellyn="Y" and oCJMall.FItemList(i).FcjmallSellYn<>"Y") or (oCJMall.FItemList(i).FSellyn<>"Y" and oCJMall.FItemList(i).FcjmallSellYn="Y") Then
	%>
					<strong><%= oCJMall.FItemList(i).FcjmallSellYn %></strong>
	<%
				Else
					response.write oCJMall.FItemList(i).FcjmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oCJMall.FItemList(i).FcjmallPrdNo)) Then
			Response.Write "<a target='_blank' href='http://www.oCJMall.com/prd/detail_cate.jsp?item_cd="&oCJMall.FItemList(i).FcjmallPrdNo&"'>"&oCJMall.FItemList(i).FcjmallPrdNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oCJMall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oCJMall.FItemList(i).FItemID%>','0');"><%= oCJMall.FItemList(i).FoptionCnt %>:<%= oCJMall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oCJMall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).Fcddkey <> "" Then
			response.write "��Ī��("&oCJMall.FItemList(i).Finfodiv&")"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If

		If (oCJMall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& oCJMall.FItemList(i).FlastErrStr &"'>ERR:"& oCJMall.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oCJMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oCJMall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oCJMall.StartScrollPage To oCJMall.FScrollCount + oCJMall.StartScrollPage - 1 %>
		<% If i>oCJMall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oCJMall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oCJMall = nothing %>
<!-- #include virtual="/lib/adminbodytail.asp" -->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->