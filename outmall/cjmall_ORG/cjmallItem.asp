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
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp) 
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'CJMall 상품코드 엔터키로 검색되게
If cjmallitemid<>"" then
	Dim iA2, arrTemp2, arrcjmallitemid
	cjmallitemid = replace(cjmallitemid,",",chr(10))
	cjmallitemid = replace(cjmallitemid,chr(13),"")
	arrTemp2 = Split(cjmallitemid,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2) 
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
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

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oCJMall.getCjmallreqExpireItemList
Else
	oCJMall.getCjmallRegedItemList			'그 외 리스트
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

//등록여부 조건 Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
}
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("<%=manageUrl%>/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=cjmall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('<%=manageUrl%>/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=cjmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM = window.open("/outmall/cjmall/popcjmallCateList.asp","popCateMancjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//상품분류 관리
function pop_prdDivManager() {
	var pCM2 = window.open("/outmall/cjmall/popcjmallprdDivList.asp","popprdDivcjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//상품 등록
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}
	if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "RegSelect";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//상품 조회
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}
	document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItem";
	document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}
//상품 상태 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="판매중";break;
		case "N": strSell="일시중단";break;
		case "X": strSell="판매종료(삭제)";break;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※cjmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 cjmall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }

		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//정보 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품 정보를 수정 하시겠습니까?\n\n※옵션추가 및 상품 정보가 수정됩니다. 가격 및 상태는 수정되지 않습니다')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//단품 판매상태 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (confirm('CjMall에 선택하신 ' + chkSel + '개 단품 상태를 일괄 수정 하시겠습니까?\n\n※CjMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditDanpum").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdSaleDTSel";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}
//단품 수량 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 단품 수량을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditqty").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditQty";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//상품 가격 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품 가격을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect2";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//단품 가격 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 단품 가격을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
       // document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}
//공통코드 검색
function popCjCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	xLink.location.href = "/outmall/cjmall/actCjMallReq.asp?cmdparam=cjmallCommonCode&CommCD="+ccd+"";
}
//정보+단품 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		//document.getElementById("btnEditSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSelect2";
		document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
//선택상품 승인확인 및 판매상태 check - batch
function batchStatCheck(){
    document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItemAuto";
	document.frmSvArr.action = "/outmall/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/outmall/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=cjmall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

function CjSellynSubmit(yn){

	if(yn == true){
	    if (document.getElementById('theday').value.length!=10) {
	        alert('날짜 형식으로 입력해 주세요.yyyy-mm-dd');
	        return;
	     }
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST&sday="+document.getElementById('theday').value+"";
	}else{
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday="+document.getElementById('sday').value+"";
	}
	document.getElementById("btnSell1").disabled=true;
	document.getElementById("btnSell2").disabled=true;
}
//이하 링크를 scm으로 보내는 이유 : 이노디터 에디터의 라이센스가 wapi.10x10.co.kr에는 없어서 작동안하기 때문
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('http://scm.10x10.co.kr/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;<% OutmallAdminInfo("cjmall") %>
		&nbsp;
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		cjmall 상품코드 : <textarea rows="2" cols="20" name="cjmallitemid" id="itemid"><%=replace(cjmallitemid,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/lib/module/categoryselectbox.asp"-->
		<br>
		등록여부 : 
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >CJmall 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >CJmall 등록예정이상
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >CJmall 등록예정
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >CJmall 전송시도중오류
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >CJmall 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >CJmall 등록완료(전시)
		</select>&nbsp;
		<label><input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">품절처리요망</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">등록상품 판매가능</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">등록여부조건Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/outmall/incsearch1.asp"-->
		카테고리
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
		상품분류
		<select name="MatchPrddiv" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>cjmall 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="cjmallYes10x10No" <%= ChkIIF(cjmallYes10x10No="on","checked","") %> ><font color=red>cjmall판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="cjmallNo10x10Yes" <%= ChkIIF(cjmallNo10x10Yes="on","checked","") %> ><font color=red>cjmall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
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
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">
			</td>
			<td align="right">
				<font color="RED">우측 2개 선작업 필요! :</font>
				<input class="button" type="button" value="cjMall상품분류매칭" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="cjMall카테고리매칭" onclick="pop_CateManager();">&nbsp;&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
<!--	<center><strong>옵션추가금액 있는 상품의 가격수정 법 :</strong> <font color="red">상품 가격 수정</font> 완료 후  <font color="red">단품 수정(가격)</font> 수정</center> -->
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
	    		실제상품 등록 :
				<input class="button" type="button" id="btnRegSel" value="상품 등록" onClick="CjSelectRegProcess(true);" <%= Chkiif(isReged = "A", "disabled", "") %> >&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSel" value="정보 수정" onClick="CjSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDanpum" value="단품 수정(상태)" onClick="CjSelectSaleStatEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditqty" value="단품 수정(수량)" onClick="CjSelectQTYEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="단품 수정(가격)" onClick="CjSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDate" value="정보+단품 수정" onClick="CjSelectEdit2Process();">
				<!--
				<input class="button" type="button" id="btnEditDate" value="선택상품 예약수정" onClick="CjSelectDateEditProcess();">
				-->
				<br><br>
				승인여부 검색 :
				<!--
				<input type="text" name="theday" value="" size="10" maxlength="10">
				<input class="button" type="button" id="btnSell1" value="특정날짜 승인여부 확인" onClick="CjSellynSubmit(true);">&nbsp;&nbsp;
				<select name="sday" class="select" id="sday">
				<% For i = 0 to 9 %>
					<option value="<%=i%>"><%=i%>
				<% Next %>
				</select>일전&nbsp;
				<input class="button" type="button" id="btnSell2" value="일정 기간 승인여부 확인" onClick="CjSellynSubmit(false);" >
				-->
				<input class="button" type="button" id="btnSelectDate" value="상품 조회" onClick="checkCjItemConfirm(this);" >
				<br><br>
				공통코드 검색 :
				<select name="CommCD" class="select" id="CommCD">
					<option value="L126">택배사코드
					<option value="6009">리드타임
					<option value="8047">가등록채널구분
				</select>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="popCjCommCDSubmit();" >
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">일시중단</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="CjmallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input class="button" type="button" value="판매상태Check(스케줄)" onClick="batchStatCheck();">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= FormatNumber(oCJMall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCJMall.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">CJMall등록일<br>CJMall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">CJMall<br>가격및판매</td>
	<td width="70">CJMall<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="60">상품분류<br>매칭여부</td>
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
			<font color="red">품절</font>
	<%
			Else
	%>
			<font color="red">일시<br>품절</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).FItemdiv = "06" OR oCJMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
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
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCJMall.FItemList(i).Fcddkey <> "" Then
			response.write "매칭됨("&oCJMall.FItemList(i).Finfodiv&")"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
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