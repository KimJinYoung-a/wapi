//###########################################################
// Description : 온라인 & 오프라인 어드민 장바구니 JS
// Hieditor : 2011.08.02 한용민 생성
//###########################################################

//장바구니 상품추가	//onoffgubun ON:온라인 OFF;오프라인
function adminshoppingbagreg(upfrm,onoffgubun,shopid){
    var frm;
    var pass = false;
    var searchfrm = document.frm;
    
    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;
    
    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (!IsInteger(frm.itemno.value)){
                    alert('갯수는 정수만 가능합니다.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value==""){
                    alert('수량을 입력하세요.');
                    frm.itemno.focus();
                    return;
                }

                //if (frm.itemno.value=="0"){
                //    alert('0이 아닌 수량을 입력하세요.');
                //    frm.itemno.focus();
                //    return;
                //}
                                                
                upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + ",";
                upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
                upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + ",";
                upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + ",";
            }
        }
    }

    //상품내역을 장바구니 페이지로 넘김..  장바구니 담는 액션은 장바구니 페이지 내에서 처리됨..
    var popbag = window.open('','popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
    upfrm.action='/common/item/adminshoppingbag.asp';
    upfrm.target='popbag';
    upfrm.onoffgubun.value=onoffgubun;
    
	//오프라인의 경우에만 매장정보 넣음
    if (onoffgubun=='OFF'){
    	upfrm.shopid.value=shopid;
    }

    upfrm.submit();
    popbag.focus();
	
	//현재 페이지 상품 선택한 내역 모두 리셋
	upfrm.itemgubunarr.value = '';
	upfrm.itemidarr.value = '';
	upfrm.itemoptionarr.value = '';
	upfrm.itemnoarr.value = '';
	upfrm.onoffgubun.value = '';
	
    if (onoffgubun=='OFF'){
		upfrm.shopid.value = '';
    }
    
    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

//장바구니 단품상품추가	//onoffgubun ON:온라인 OFF;오프라인
function adminshoppingbagregoneitem(onoffgubun,shopid,upfrm){
    if (onoffgubun=="" || upfrm.itemgubun.value=="" || upfrm.itemgubun.value=="" || upfrm.itemid.value=="" || upfrm.itemoption.value==""){
        alert('값이 없습니다');
        upfrm.itemno.focus();
        return;
    }
	
	//오프라인의 경우
    if (onoffgubun=='OFF'){
    	//매장정보 넣음
	    if (shopid==""){
	        alert('매장이 없습니다');
	        return;
	    }
    }    
    
    if (!IsInteger(upfrm.itemno.value)){
        alert('갯수는 정수만 가능합니다.');
        upfrm.itemno.focus();
        return;
    }

    if (upfrm.itemno.value==""){
        alert('수량을 입력하세요.');
        upfrm.itemno.focus();
        return;
    }
    
    var tmp = '&itemgubunarr='+upfrm.itemgubun.value+',&itemidarr='+upfrm.itemid.value+',&itemoptionarr='+upfrm.itemoption.value+',&itemnoarr='+upfrm.itemno.value+',';
    var popbag = window.open('/common/item/adminshoppingbag_process.asp?mode=directbagaddarr&onoffgubun='+onoffgubun+'&shopid='+shopid+tmp,'popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
	popbag.focus();
}

//장바구니 보기	//onoffgubun ON:온라인 OFF;오프라인
function adminshoppingbagview(upfrm,onoffgubun,shopid){

    var popbag = window.open('','popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
    upfrm.onoffgubun.value=onoffgubun;

    if (onoffgubun=='OFF'){
    	//upfrm.shopid.value=shopid;
    }
    
    upfrm.action='/common/item/adminshoppingbag.asp';
    upfrm.target='popbag';
    upfrm.submit();
    popbag.focus();
}

//필요수량클릭시	선택한 필요수량을 수량으로 넣는다 
function inputiteno(shortitemno,formi){
    formi.itemno.value=shortitemno;

    formi.cksel.checked=true;
    AnCheckClick(formi.cksel);  
}

//검색버튼  
function reg(upfrm){

    if(upfrm.itemid.value!=''){
        if (!IsDouble(upfrm.itemid.value)){
            alert('상품코드는 숫자만 가능합니다.');
            upfrm.itemid.focus();
            return;
        }
    }

    upfrm.submit();
}

//장바구니 수정
function bageditarr(upfrm){
    var frm;
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (!IsInteger(frm.itemno.value)){
                    alert('갯수는 정수만 가능합니다.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('수량을 입력하세요.');
                    frm.itemno.focus();
                    return;
                }
				
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + ",";
                upfrm.bagidxarr.value = upfrm.bagidxarr.value + frm.bagidx.value + ",";
            }
        }
    }
	
    upfrm.action='/common/item/adminshoppingbag_process.asp';
    upfrm.mode.value='bageditarr';
    upfrm.target='view';
    upfrm.submit();
}

//장바구니 상품 삭제
function bagdelarr(upfrm){
    var frm;
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                upfrm.bagidxarr.value = upfrm.bagidxarr.value + frm.bagidx.value + ",";
            }
        }
    }
    
    upfrm.action='/common/item/adminshoppingbag_process.asp';
    upfrm.mode.value='bagdelarr';
    upfrm.target='view';
    upfrm.submit();
}

//텐바이텐물류 주문서 작성
function AddArr(upfrm ,shopgubun){
    var frm; var tmpshopid = ''; var tmpcomm_cd013 = ''; var tmpcomm_cd011 = ''; var tmpcomm_cd031 = '';
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.buycasharr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";
    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                
                //텐바이텐물류 주문의 경우 주문자가 동일해야함
                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('서로 틀린 매장이 선택되어 있습니다 \n매장(주문자)이 동일 해야 합니다.');
	                    return;
                	}
                }
				
				//텐바이텐물류 주문의 경우 출고분정산과 텐바이텐위탁만 주문가능
                if (frm.comm_cd.value=="B012" || frm.comm_cd.value=="B022"){
                    alert('업체위탁이나 업체매입은 주문 하실수 없습니다.');
                    frm.itemno.focus();
                    return;
                }

				//텐바이텐 위탁 주문
                if (frm.comm_cd.value=="B011" && tmpcomm_cd011==''){
                	tmpcomm_cd011 = frm.comm_cd.value;
                }
				//출고매입
                if (frm.comm_cd.value=="B031" && tmpcomm_cd031==''){
                	tmpcomm_cd031 = frm.comm_cd.value;
                }
				//출고위탁
                if (frm.comm_cd.value=="B013" && tmpcomm_cd013==''){
                	tmpcomm_cd013 = frm.comm_cd.value;
                }
                
                //텐바이텐물류 주문 출고위탁의 경우, 출고위탁끼리만 주문이 가능함
                if (tmpcomm_cd013 != ''){
                	
                	if (tmpcomm_cd011 != '' || tmpcomm_cd031 != ''){
                		alert('출고위탁의 경우, 텐바이텐위탁과 출고매입 주문과 같이 주문하실수 없습니다.');
                		return;
                	}
                	upfrm.cwflag.value='1';
                }else{
                	upfrm.cwflag.value='0';
                }

                if (!IsInteger(frm.itemno.value)){
                    alert('갯수는 정수만 가능합니다.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('수량을 입력하세요.');
                    frm.itemno.focus();
                    return;
                }
				
                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";
                
                //[db_storage].[dbo].tbl_ordersheet_master에 들어가는 내용의 경우 센터매입가와 , 매장매입가가 꺼꾸로임  
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopbuyprice.value + "|";
                upfrm.buycasharr2.value = upfrm.buycasharr2.value + frm.shopsuplycash.value + "|";
            	
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
				frmbag.bagidxarr.value = frmbag.bagidxarr.value + frm.bagidx.value + ",";
            
            }
        }
    }
	
	//선택한 내역 장바구니에서 삭제
    frmbag.action='/common/item/adminshoppingbag_process.asp';
    frmbag.mode.value='baginsertdelarr';
    frmbag.target='view';
    frmbag.submit();

    //텐바이텐물류 주문서 작성페이지
    //매장
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
	//직원
	}else{
		upfrm.action='/admin/fran/jumuninput.asp';
	}
    upfrm.shopid.value=tmpshopid;
    upfrm.submit();
}

//업체 주문서 작성
function AddArr_upche(upfrm,shopgubun){
    var frm; var tmpshopid = ''; var tmpmakerid = ''; var ret;
    var pass = false;
    var searchfrm = document.frm;
    
    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    //if (searchfrm.makerid.value == ''){
    //    alert('브랜드(공급처)를 선택해 주세요.');
    //    return;
    //}
    
    if (!pass) {
        alert('선택 아이템이 없습니다.');
        return;
    }
    
    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.shopbuypricearr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";
    upfrm.bagidxarr.value = "";
    
    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
				
				//업체주문의 경우 주문자가 동일해야함
                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('서로 틀린 매장이 선택되어 있습니다 \n매장(주문자)이 동일 해야 합니다.');
	                    return;
                	}
                }
				
				//업체주문의 경우 하나의 주문에 공급처는 한개여야 한다
                if (tmpmakerid==''){
                	tmpmakerid = frm.makerid.value;
                } else {
                	if (tmpmakerid != frm.makerid.value){
	                    alert('서로 틀린 브랜드(공급처)가 선택되어 있습니다 \n업체주문의 경우 브랜드(공급처)가 동일해야 합니다');
	                    return;
                	}
                }
				
				//업체주문의경우 업체위탁과 업체매입만 주문가능
                if (frm.comm_cd.value=="B011" || frm.comm_cd.value=="B031" || frm.comm_cd.value=="B013"){
                    alert('텐바이텐위탁, 출고매입, 출고위탁은 주문 하실수 없습니다.');
                    frm.itemno.focus();
                    return;
                }               
                
                if (!IsInteger(frm.itemno.value)){
                    alert('갯수는 정수만 가능합니다.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('수량을 입력하세요.');
                    frm.itemno.focus();
                    return;
                }

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.shopbuypricearr2.value = upfrm.shopbuypricearr2.value + frm.shopbuyprice.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
				frmbag.bagidxarr.value = frmbag.bagidxarr.value + frm.bagidx.value + ",";

            }
        }
    }
	
	//선택한 내역 장바구니에서 삭제
    frmbag.action='/common/item/adminshoppingbag_process.asp';
    frmbag.mode.value='baginsertdelarr';
    frmbag.target='view';
    frmbag.submit();
    
    //업체 주문서 작성 페이지	
    //매장
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
	//직원
	}else{
		upfrm.action='/common/offshop/shop_ipchulinput.asp';
	}
    upfrm.shopid.value=tmpshopid;
    upfrm.chargeid.value=tmpmakerid;
    upfrm.submit();
}

//브랜드클릭시
function searchmakerid(makerid,upfrm){
    upfrm.makerid.value=makerid;
    upfrm.submit();
}

function CheckThis(frm){
    frm.cksel.checked=true;
    AnCheckClick(frm.cksel);
}

// 새상품 추가 팝업
function addnewItem(onoffgubun,upfrm,shopid ,acURL){
	var addnewItem; var tmpshopid;
	
	tmpshopid = shopid;
	//tmpshopid = upfrm.shopid.value;
	
	if (onoffgubun=='ON'){
		
	}else if (onoffgubun=='OFF'){
		if (tmpshopid==''){
			alert('상품을 추가 하실 매장을 검색 하시고, 상품을 추가 하세요');
			upfrm.shopid.focus();
			return;
		}

		addnewItem = window.open("/common/offshop/pop_itemAddInfo_off.asp?shopid="+tmpshopid+"&acURL="+acURL, "addnewItem", "width=1024,height=768,scrollbars=yes,resizable=yes");
		addnewItem.focus();
	}
}