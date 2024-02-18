/* 계약직 급여관리 스크립트 */

// 4.345238095 == 월 평균 WEEK 수 = (365일 / 7일 / 12개월)
// 평균주수
var avgWeek = 4.345238095;

//주휴일 설정시 근무시간 설정 폼 비활성화
function jsSetWH(iWd){
	if (eval("document.frmPay.selWH"+iWd).value =="3"){	//주휴일
		eval("document.frmPay.iSH"+iWd).disabled =true;
		eval("document.frmPay.iSM"+iWd).disabled =true;
		eval("document.frmPay.iEH"+iWd).disabled =true;
		eval("document.frmPay.iEM"+iWd).disabled =true;
		eval("document.frmPay.iBSH"+iWd).disabled =true;
		eval("document.frmPay.iBSM"+iWd).disabled =true;
		eval("document.frmPay.iBEH"+iWd).disabled =true;
		eval("document.frmPay.iBEM"+iWd).disabled =true;
		eval("document.frmPay.iD"+iWd).disabled =true;
		eval("document.frmPay.iSH"+iWd).value = "";
		eval("document.frmPay.iSM"+iWd).value = "";
		eval("document.frmPay.iEH"+iWd).value = "";
		eval("document.frmPay.iEM"+iWd).value = "";
		eval("document.frmPay.iBSH"+iWd).value = "";
		eval("document.frmPay.iBSM"+iWd).value = "";
		eval("document.frmPay.iBEH"+iWd).value = "";
		eval("document.frmPay.iBEM"+iWd).value = "";
		eval("document.frmPay.iD"+iWd).value = "";
		eval("document.frmPay.intd"+iWd).value = "";

	}else{	//근무일
		eval("document.frmPay.iSH"+iWd).disabled =false;
		eval("document.frmPay.iSM"+iWd).disabled =false;
		eval("document.frmPay.iEH"+iWd).disabled =false;
		eval("document.frmPay.iEM"+iWd).disabled =false;
		eval("document.frmPay.iBSH"+iWd).disabled =false;
		eval("document.frmPay.iBSM"+iWd).disabled =false;
		eval("document.frmPay.iBEH"+iWd).disabled =false;
		eval("document.frmPay.iBEM"+iWd).disabled =false;
		eval("document.frmPay.iD"+iWd).disabled =false;
		eval("document.frmPay.iWHT"+iWd).value = "";
		document.all.totWHT.innerHTML =  "";
	}

	jsSetWeekHoliday();
}

// 근무 종료시간-시작시간 - 휴계시간 = 총근무시간
function jsCalDutyTime(iWd){
	jsSetDutyTime(iWd);
	jsSetWeekHoliday();
	jsSetMonthlypay();
}

//근무시간 계산
function jsSetDutyTime(iWd){
	var istarthour,istartminute,iendhour,iendminute,ibreakhour,ibreakminute;
	var iduty;
	var inBreakTime = "0";
	var inighttime = 0;

	if(document.frmPay.blnBT.checked){inBreakTime = "1";}
	istarthour = eval("document.frmPay.iSH"+iWd).value;
	istartminute =  eval("document.frmPay.iSM"+iWd).value;
	iendhour = eval("document.frmPay.iEH"+iWd).value;
	iendminute = eval("document.frmPay.iEM"+iWd).value;
	ibreakshour = eval("document.frmPay.iBSH"+iWd).value;
	ibreaksminute = eval("document.frmPay.iBSM"+iWd).value;
	ibreakehour = eval("document.frmPay.iBEH"+iWd).value;
	ibreakeminute = eval("document.frmPay.iBEM"+iWd).value;

	if (istarthour =="" ){istarthour=0;}
	if (istartminute =="" ){istartminute =0}
	if (iendhour =="" ){iendhour =0}
	if (iendminute=="" ){iendminute=0}
	if (ibreakshour =="" ){ibreakshour =0}
	if (ibreaksminute =="" ){ibreaksminute =0}
	if (ibreakehour =="" ){ibreakehour =0}
	if (ibreakeminute =="" ){ibreakeminute =0}

	istarthour = parseInt(istarthour,10);
	istartminute =  parseInt(istartminute,10);
	iendhour = parseInt(iendhour,10);
	iendminute =  parseInt(iendminute,10);
	ibreakshour =  parseInt(ibreakshour,10);
	ibreaksminute =  parseInt(ibreaksminute,10);
	ibreakehour =  parseInt(ibreakehour,10);
	ibreakeminute =  parseInt(ibreakeminute,10);

	//근무시간
	iduty = (iendhour*60+ iendminute)-  (istarthour*60+ istartminute);
	var ibreak =(ibreakehour*60+ibreakeminute)-(ibreakshour*60+ibreaksminute);

	//휴계시간 포함여부
	if(inBreakTime=="0") {iduty = iduty- ibreak; }

	var nightS, nightE, nightBS, nightBE;
	//야간근무수당
	if((iendhour*60+ iendminute)>22*60){
		if ((istarthour*60+ istartminute) < 22*60){
			nightS = 22*60;
		}else{
			nightS = istarthour*60+ istartminute;
		}

		if ((iendhour*60+ iendminute) > 30*60){
			nightE = 30*60;
		}else{
			nightE = iendhour*60+ iendminute;
		}

		if ((ibreakshour*60+ ibreaksminute) < 22*60){
			nightBS = 22*60;
		}else if((ibreakshour*60+ ibreaksminute) >=30*60){
			nightBS = 0;
		}else{
			nightBS = ibreakshour*60+ ibreaksminute;
		}

		if ((ibreakehour*60+ibreakeminute) < 22*60){
			nightBE = 22*60;
		}else if((ibreakehour*60+ ibreakeminute) >30*60){
			nightBE = 0;
		}else{
			nightBE = ibreakehour*60+ibreakeminute;
		}

		if(inBreakTime=="0"){
			inighttime = nightE- nightS- (nightBE-nightBS);
		}else{
			inighttime = nightE- nightS;
		}
	}

	eval("document.frmPay.iD"+iWd).value =jsTimeForm(iduty);
	eval("document.frmPay.intd"+iWd).value =inighttime;
}

//총 시간 설정
function jsSetWeekHoliday() {
	var arrValue, iValue, iNightTime, arrTime;
	var totDuty  =0;
	var iWHD = 0;
	var totNightTime = 0;
	var realWeekWorkDay = 0;

	for(i=1;i<8;i++){
		if( eval("document.frmPay.selWH"+i).value=="3"){iWHD=i};

		if(eval("document.frmPay.iD"+i).value==""){
			iValue = 0;
		}else{
			arrValue = eval("document.frmPay.iD"+i).value.split(":");
			iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
		}
		totDuty = totDuty + iValue;

		if((eval("document.frmPay.selWH"+i).value == "1") && (iValue >= 60)) {
			// 한시간 이상 일했으면 근무일수에 산정(식대산정에 사용한다.)
			realWeekWorkDay = realWeekWorkDay + 1;
		}

		iNightTime =eval("document.frmPay.intd"+i).value;
		if (iNightTime==""){iNightTime=0;}
		totNightTime = totNightTime + parseInt(iNightTime,10);
	}

	document.all.totDuty.innerHTML =  jsTimeForm(totDuty);

	var totDutyH = parseInt(totDuty/60,10);

	// 주휴시간을 분단위로 계산
	var WHTime = 0;
	if(totDutyH>=15){
		// WHTime = (8*(totDutyH/40))*60;
		WHTime = (8*(totDuty/40));
	}

	if (totDutyH>40){
		// WHTime=	(8*(40/40))*60;
		WHTime = 480;
	}
	//alert(totDutyH);
	if (iWHD > 0){ eval("document.frmPay.iWHT"+iWHD).value = jsTimeForm(WHTime);}
	if (totDutyH>=15 && iWHD > 0){document.all.totWHT.innerHTML =  jsTimeForm(WHTime);}

	document.frmPay.totnt.value = Math.ceil(totNightTime/60*avgWeek) ;
	document.frmPay.totdt.value =  Math.ceil(totDuty/60*avgWeek);
	document.frmPay.totwhdt.value = Math.ceil(WHTime/60*avgWeek);
	document.frmPay.totot.value =  document.frmPay.iot.value;
	document.frmPay.totd.value =  Math.ceil(realWeekWorkDay * avgWeek);
}

//휴계시간 근무시간에 포함여부에 따른 근무시간, 주휴시간, total 재조정
function jsSetInBreakTime(){
	for(i=1;i<8;i++){
		jsSetDutyTime(i);
	}

	jsSetWeekHoliday();
	jsSetMonthlypay();
}

//자동 필드넘김
function TnTabNumber(thisform,target,num) {
	if (eval("document.frmPay." + thisform + ".value.length") == num) {
		if(!eval("document.frmPay." + target + ".disabled")){
			eval("document.frmPay." + target + ".focus()");
			eval("document.frmPay." + target + ".select()");
		}
	}
}

//시간 폼 맞춤(시간,분을 두자리 숫자로 ex:01:03)
function jsTimeForm(totMin){
	var iHour = parseInt(totMin/60,10);
	var iMinute = totMin%60;

	if(String(iHour).length < 2){
		iHour ="0"+iHour;
	}
	if(String(iMinute).length < 2){
		iMinute ="0"+iMinute;
	}
	return iHour+":"+iMinute;
}

//시간 폼에서 분으로 재계산(01:30 폼에서 -> 90 으로)
function jsFormToTime(strForm){
	var arrValue = strForm.split(":");
	return parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
}


//-- jsPopCal : 달력 팝업 --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}


//월급여합계 설정
function jsSetMonthlypay(){
	var mdefaultpay =  document.frmPay.iHP.value;
	var mdefaultfoodpay =  document.frmPay.iEP.value;

	document.frmPay.idp.value =  mdefaultpay*document.frmPay.totdt.value;
	document.frmPay.ifp.value =  mdefaultfoodpay*document.frmPay.totd.value;
	document.frmPay.iwhdp.value = mdefaultpay* document.frmPay.totwhdt.value;
	document.frmPay.iotp.value = mdefaultpay*document.frmPay.iot.value*1.5;
	document.frmPay.inp.value = mdefaultpay*document.frmPay.totnt.value*0.5;

	document.frmPay.itotp.value = parseInt(document.frmPay.idp.value,10) + parseInt(document.frmPay.ifp.value,10) + parseInt(document.frmPay.iwhdp.value,10) + parseInt(document.frmPay.iotp.value,10) + parseInt(document.frmPay.inp.value,10);
}

//시간외 수당변경 화면 보여주기
function jsSetOverTime(){
	if(document.frmPay.blnOT.checked){
		document.all.spanOT.style.display ="";
		document.frmPay.iot.value ="22";
		document.frmPay.totot.value ="22";
		document.frmPay.iotp.value = document.frmPay.iHP.value*document.frmPay.iot.value*1.5;
	}else{
		document.all.spanOT.style.display ="none";
		document.frmPay.iot.value ="0";
		document.frmPay.totot.value ="0";
		document.frmPay.iotp.value ="0";
	}
	document.frmPay.itotp.value = parseInt(document.frmPay.idp.value,10)+parseInt(document.frmPay.iwhdp.value,10)+parseInt(document.frmPay.iotp.value,10)+parseInt(document.frmPay.inp.value,10);
}


//시간외수당변경에 따른 급여합계 변경
function jsSetOverTimePay(){
	if(document.frmPay.blnOT.checked){
		document.frmPay.iotp.value =document.frmPay.iHP.value*document.frmPay.iot.value*1.5;
	}
	document.frmPay.itotp.value = parseInt(document.frmPay.idp.value,10)+parseInt(document.frmPay.iwhdp.value,10)+parseInt(document.frmPay.iotp.value,10)+parseInt(document.frmPay.inp.value,10);
}

// 기본 계약상의 주휴시간
// 유급휴일이면 근무시간을 기본 계약상의 주휴시간으로 설정한다.
var holidaywdtime = 0;
function jsSetHolidayWD(v) {
	holidaywdtime = v; 
}

//근무시간 입력에 따른 각각 총합계산(pop_workitem.asp)
function jsSetTotTime(iWd){ 
	var istarthour,istartminute,iendhour,iendminute,ibreakhour,ibreakminute, iouthour, ioutminute;
	var iduty;
	var inBreakTime = "False";
	var inighttime = 0;
	var iextendtime = 0;
	var iholidaytime =0;
	var iwholidaytime =0;

	if (eval("document.frmPay.selWH"+iWd) == undefined) {
		return;
	}

	inBreakTime = document.frmPay.hidInB.value;

	istarthour = eval("document.frmPay.iSH"+iWd).value;
	istartminute =  eval("document.frmPay.iSM"+iWd).value;
	iendhour = eval("document.frmPay.iEH"+iWd).value;
	iendminute = eval("document.frmPay.iEM"+iWd).value;
	ibreakshour = eval("document.frmPay.iBSH"+iWd).value;
	ibreaksminute = eval("document.frmPay.iBSM"+iWd).value;
	ibreakehour = eval("document.frmPay.iBEH"+iWd).value;
	ibreakeminute = eval("document.frmPay.iBEM"+iWd).value;
	iouthour	= eval("document.frmPay.iOH"+iWd).value;
	ioutminute	= eval("document.frmPay.iOM"+iWd).value;

	if (istarthour =="" ){istarthour=0;}
	if (istartminute =="" ){istartminute =0}
	if (iendhour =="" ){iendhour =0}
	if (iendminute=="" ){iendminute=0}
	if (ibreakshour =="" ){ibreakshour =0}
	if (ibreaksminute =="" ){ibreaksminute =0}
	if (ibreakehour =="" ){ibreakehour =0}
	if (ibreakeminute =="" ){ibreakeminute =0}
	if (iouthour =="" ){iouthour =0}
	if (ioutminute =="" ){ioutminute =0}

	istarthour = parseInt(istarthour,10);
	istartminute =  parseInt(istartminute,10);
	iendhour = parseInt(iendhour,10);
	iendminute =  parseInt(iendminute,10);
	ibreakshour =  parseInt(ibreakshour,10);
	ibreaksminute =  parseInt(ibreaksminute,10);
	ibreakehour =  parseInt(ibreakehour,10);
	ibreakeminute =  parseInt(ibreakeminute,10);
	iouthour =  parseInt(iouthour,10);
	ioutminute =  parseInt(ioutminute,10);

	//근무시간
	iduty = (iendhour*60+ iendminute)-  (istarthour*60+ istartminute) - (iouthour*60+ioutminute);
	var ibreak =(ibreakehour*60+ibreakeminute)-(ibreakshour*60+ibreaksminute);

	//휴계시간 포함여부
	if (inBreakTime=="False") {iduty = iduty - ibreak; }

	// 유급휴일  
	// 유급휴일이면 근무시간에 기본 계약상의 주휴시간으로 설정한다. 
	// => 변경.(2014-08-11 정윤정) 휴가등록시 등록한 시간이 연차 필드에 추가된다. 추가 근무 가능 
	//if (eval("document.frmPay.selWH"+iWd).value =="4") {
		// eval("document.frmPay.iwhWT"+iWd).value =jsTimeForm(holidaywdtime);
	//	iduty = holidaywdtime;
	//}

	//연장근무
	//		 	var idayS = iWd-(parseInt(eval("document.frmPay.hidWD"+iWd).value,10)-1) ;
	//	 		var chkExt = 0;
	//		 	var iEExt = 0;
	//		 	var totEExt = 0;
	//	 		var arrExt;
	//
	//	 		if (idayS >0){ //일요일(idays)부터 해당요일(iwd)이전까지 근무시간 합 계산
	//	 		 for(i=idayS;i<iWd;i++){
	//		  		arrExt = eval("document.frmPay.iWT"+i).value.split(":");
	//	 	  		iEExt = parseInt(arrExt[0],10)*60+parseInt(arrExt[1],10);
	//		 	 		totEExt = totEExt + iEExt;
	//		 	 }
	//	 		}
	//
  
	if (iduty>480){//하루 8시간 이상근무일때 연장근무
		iextendtime = iduty - 480;
		iduty = 480;
	} 
	//
	//		 	if (totEExt>=2400){//1.근무시간 합이 40시간 이상일때 연장근무처리
	//		 		iextendtime = iextendtime+iduty;
	//		 		iduty = 0;
	//		 	}else if((totEExt+iduty)>2400){//2.현재날짜와 전날까지의 근무시간 합이 40시간 이상일때 연장근무처리
	//		 		iextendtime = iextendtime+(2400-totEExt);
	//		 		iduty = iduty-(2400-totEExt);
	//		 	}

	var nightS, nightE, nightBS, nightBE;
	//야간근무수당

	if((iendhour*60+ iendminute)>22*60){  
		if ((istarthour*60+ istartminute) < 22*60){
			nightS = 22*60;
		}else{
			nightS = istarthour*60+ istartminute;
		}

		if ((iendhour*60+ iendminute) > 30*60){
			nightE = 30*60;
		}else{
			nightE = iendhour*60+ iendminute;
		}

		if ((ibreakshour*60+ ibreaksminute) < 22*60){
			nightBS = 22*60;
		}else if((ibreakshour*60+ ibreaksminute) >=30*60){
			nightBS = 0;
		}else{
			nightBS = ibreakshour*60+ ibreaksminute;
		}

		if ((ibreakehour*60+ibreakeminute) < 22*60){
			nightBE = 22*60;
		}else if((ibreakehour*60+ ibreakeminute) >30*60){
			nightBE = 0;
		}else{
			nightBE = ibreakehour*60+ibreakeminute;
		}

		// if(inBreakTime=="0"){
		if (inBreakTime=="False") {
			inighttime = nightE - nightS - (nightBE - nightBS);
		}else{
			inighttime = nightE - nightS;
		}
	}


	//휴일근무수당
	if(eval("document.frmPay.selWH"+iWd).value =="3" || eval("document.frmPay.selWH"+iWd).value =="5"){ 
		if(iduty>0){
			iholidaytime = iduty+iextendtime;
			//	eval("document.frmPay.iwhWT"+iWd).value = jsTimeForm(0);
		}else{ 
			jsChangeWeekHoliday_Pre(eval("document.frmPay.hidWD"+iWd).value,iWd);
		}
	}

	if ( iduty < 0 ){
		iduty = 0;
		iextendtime = 0;
		inighttime = 0;
		iholidaytime = 0;
	}

	eval("document.frmPay.iWT"+iWd).value =jsTimeForm(iduty);
	eval("document.frmPay.ieWT"+iWd).value =jsTimeForm(iextendtime);
	eval("document.frmPay.inWT"+iWd).value =jsTimeForm(inighttime);
	eval("document.frmPay.ihWT"+iWd).value =jsTimeForm(iholidaytime);
	 jsChangeWeekHoliday(eval("document.frmPay.hidWD"+iWd).value,iWd);

	jsSetTotTimeSum();  
}


//주휴시간 설정
function jsChangeWeekHoliday(iWeekday,iWorkday){
	var arrValue, iValue, iEValue, iVValue, iNightTime, arrTime, iEday, iSday, chkWHD;
	var totDuty  =0;
	var totWorkDuty = 0;
	var iWorkD;
	var iWHD = -1;
	var totNightTime = 0;
	var chkWHD = 0;


	//현재 날짜 포함 일주일 시작일 종료일 계산
	//이전 일주일 포함 계산
	var iSWday = parseInt(iWorkday,10)-parseInt(iWeekday,10)+1;
	var iEWday = iSWday+6;

	if(iSWday < 0){
		iSWday = parseInt(document.frmPay.hidPED.value,10)+iSWday+1;
		if (iSWday < 0){ iSWday = 0}
		if(iEWday< 0){
			iEWday = parseInt(document.frmPay.hidPED.value,10)+iEWday;
		}else{
			iEWday = parseInt(document.frmPay.hidPED.value,10);
		}
  		 
		for(i=iSWday;i<=iEWday;i++){
			iValue = 0;
			iEValue = 0;
			iVValue = 0;
			arrValue = "";
			if(typeof(eval("document.frmPay.iPWT"+i)) !="undefined"){
				if(eval("document.frmPay.iPWT"+i).value!=""){
					arrValue = eval("document.frmPay.iPWT"+i).value.split(":");
					iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

					arrValue = eval("document.frmPay.iPeWT"+i).value.split(":");
					iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
					
					arrValue = eval("document.frmPay.iPVT"+i).value.split(":");
					iVValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
				}

				if( eval("document.frmPay.iPWH"+i).value=="1" && ( iValue+iEValue) == 0 ){
					chkWHD = chkWHD + 1 ;
				}
				totDuty = totDuty + iValue+iEValue+iVValue;
					 
			}
		}
	}

   
	//다시 재 설정	 - 현재 달 계산 
	var iSWday = parseInt(iWorkday,10)-parseInt(iWeekday,10)+1 ;
	var iEWday = iSWday+6;
	var iCday=0;

	var pWDay=0;            //2013/01/02 추가
	var iPSWday, iPEWday;   //2013/01/02 추가 
	if (iSWday< 0 ){
		pWDay = iSWday      //2013/01/02 추가
		iSWday=0
	}
	
//	if(iSWday < document.frmPay.hidSday.value){iSWday=document.frmPay.hidSday.value}
	if(iEWday > document.frmPay.hidEday.value){iEWday=document.frmPay.hidEday.value}
  if(iEWday< 0){	iEWday = 0;		}
//	//2013/01/02 추가
//	if (pWDay<0&&document.frmPay.hidPED.value>0){//이전 달 내용이 없으면(이전달 말일이 0이면 ) 이전달 일주일 계산 하지 않는다.
//		//이전달근무 계산
//		iPSWday = parseInt(document.frmPay.hidPED.value,10)+pWDay+1;
//		if(iEWday<0){
//			iPEWday = parseInt(document.frmPay.hidPED.value,10)+iEWday+1;
//		}else{
//			iPEWday = parseInt(document.frmPay.hidPED.value,10);
//		}
// 
//		for(i=iPSWday;i<=iPEWday;i++){ //수정된 날짜가 포함된 일주일
//			iValue = 0;
//			iEValue = 0;
//			arrValue = "";
//
//			if(typeof(eval("document.frmPay.iPWT"+i)) !="undefined"){
//				if(eval("document.frmPay.iPWT"+i).value!=""){
//					arrValue = eval("document.frmPay.iPWT"+i).value.split(":");
//					iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
//
//					//arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
//					//iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);  //??
//				}
//			}
//			totWorkDuty = totWorkDuty + iValue; //기본근무시간만 
//	}
//} 
// alert( iSWday+"-"+iEWday);
	for(i=iSWday;i<=iEWday;i++){ //수정된 날짜가 포함된 일주일
		iValue = 0;
		iEValue = 0;
		iVValue = 0;
		arrValue = "";

		if(eval("document.frmPay.iWT"+i).value!=""){
			arrValue = eval("document.frmPay.iWT"+i).value.split(":");
			iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

			arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
			iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			
			arrValue = eval("document.frmPay.iVT"+i).value.split(":");
			iVValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
		}

		if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){
			chkWHD = chkWHD + 1 ;
		}

	// 	if(i ==0 &&  document.frmPay.hidWD0.value==1){ //이전달 계산없이 주휴일로 데이터 시작할 경우만 이전달값 합해준다  ? 언제 필요한거지?2014.02.11삭제
	 //		totDuty = totDuty + iValue+iEValue + parseInt(document.frmPay.hidPWD.value,10); 
	// 	}else{
			totDuty = totDuty + iValue+iEValue+iVValue;
	// 	}
 
//		//기본근무시간+연장근무시간
//		totWorkDuty = totWorkDuty + iValue; //기본근무시간만
 
		//연장근무수당 
		if  (eval("document.frmPay.selWH"+i).value !="3" && iValue > 0 ){
			if(totDuty>2400){ 
				iCday = totDuty-2400; 
				icalWT = (iValue+iEValue)-iCday;  
				if(icalWT>480){
					icalWT = 480;
				}else if(icalWT <0){
					icalWT= 0;
				} 
				icaleWT = (iValue+iEValue)- icalWT;  
				eval("document.frmPay.ieWT"+i).value = jsTimeForm(icaleWT);
				eval("document.frmPay.iWT"+i).value = jsTimeForm(icalWT);  
			}  
//			if  ((totDuty-iValue)>=2400) {
//				eval("document.frmPay.ieWT"+i).value = jsTimeForm(iEValue+iValue);
//				eval("document.frmPay.iWT"+i).value = "00:00";
//			}else if(totDuty>2400){
//				iCday = totDuty-2400;
//				eval("document.frmPay.ieWT"+i).value = jsTimeForm(iEValue+iCday);
//				eval("document.frmPay.iWT"+i).value = jsTimeForm(iValue-(iCday)); 
//			}
		}
	}
	
	iSday  = parseInt(iWorkday,10)- parseInt(iWeekday,10)+8 ;
	iEday = iSday+6;
	
	if (iSday<= 0 ){ 	iSday=0; 	}
	if(iSday > document.frmPay.hidEday.value){return;}
	if(iEday > document.frmPay.hidEday.value){iEday=document.frmPay.hidEday.value}
  if(iEday< 0){	iEday = 0;		} 
 
	for(i=iSday;i<=iEday;i++){	 //수정된 날짜가 포함된 일주일의 그 다음주 주휴일에 값 변경된다.
		if( eval("document.frmPay.selWH"+i).value=="3"){iWHD=i};
	}
 
	var totDutyH = parseInt(totDuty/60,10);
 
	//주휴수당
	var WHTime = 0;
	if (totDutyH > 40){
		totDutyH=40;
	}
	if(totDutyH>=15){WHTime= (8*(totDutyH/40))*60;}
	 
	if (iWHD >=0 ){
		if (totDutyH>=15 && iWHD >= 0 && chkWHD ==0){
			eval("document.frmPay.iwhWT"+iWHD).value = jsTimeForm(WHTime);
		}else{
			eval("document.frmPay.iwhWT"+iWHD).value = jsTimeForm(0);
		}
	}


}

//주휴시간 설정
function jsChangeWeekHoliday_Pre(iWeekday,iWorkday){
	var arrValue, iValue, iVValue, iEValue,  iNightTime, arrTime, iEday, iSday;
	var totDuty  =0;
	var totWorkDuty =0;
	var iWHD = 0;
	var totNightTime = 0;
	var chkWHD = 0;

	if( eval("document.frmPay.selWH"+iWorkday).value !="3"){
		eval("document.frmPay.iwhWT"+iWorkday).value = jsTimeForm(0);
	}else{
		iWHD = iWorkday;
		//현재 날짜 이전 일주일 시작일 종료일 계산
		var iSWday = parseInt(iWorkday,10)-parseInt(iWeekday,10)+1-7 ;
		var iEWday = iSWday+6;
 
		if(iSWday<0){ 
			iSWday = parseInt(document.frmPay.hidPED.value,10)+iSWday+1;
 			if (iSWday < 0){
 				iSWday = document.frmPay.hidPSD.value;
 			}
			if(iEWday< 0){
 				iEWday = parseInt(document.frmPay.hidPED.value,10)+iEWday+1;
 			}else{
				iEWday = parseInt(document.frmPay.hidPED.value,10);
			}
 //alert(iSWday+"-"+iEWday);
			for(i=iSWday;i<=iEWday;i++){
				iValue = 0;
				iVValue = 0;
				iEValue = 0;
				arrValue = "";
				if(typeof(eval("document.frmPay.iPWT"+i)) !="undefined"){
					if(eval("document.frmPay.iPWT"+i).value!=""){
						//기본근무
						arrValue = eval("document.frmPay.iPWT"+i).value.split(":");
						iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
 						
 						//연장근무
						arrValue = eval("document.frmPay.iPeWT"+i).value.split(":");
						iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
						
						//유급휴가
						arrValue = eval("document.frmPay.iPVT"+i).value.split(":");
						iVValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
					}

					if( eval("document.frmPay.iPWH"+i).value=="1" && ( iValue+iEValue) == 0 ){
						chkWHD = chkWHD + 1 ;
					}
					totDuty = totDuty + iValue+iEValue+iVValue;
				}
			}
		}
	 
		//다시 재 설정
		iSWday = parseInt(iWorkday,10)-parseInt(iWeekday,10)+1-7 ;
		iEWday = iSWday+6;
		if(iSWday<0){
			iSWday=0;
		}
		if(iEWday > document.frmPay.hidEday.value){iEWday=document.frmPay.hidEday.value}
		if(iEWday< 0){ 
			iEWday =-1;
			}
  // alert(iSWday+"-"+iEWday);
		for(i=iSWday;i<=iEWday;i++){
			iValue = 0;
			iEValue = 0;
			iVValue = 0;
			arrValue = "";
			if(eval("document.frmPay.iWT"+i).value!=""){
				arrValue = eval("document.frmPay.iWT"+i).value.split(":");
				iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

				arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
				iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
				
				//유급휴가
				arrValue = eval("document.frmPay.iVT"+i).value.split(":");
				iVValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			}

			if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){
				chkWHD = chkWHD + 1 ;
			}
			totDuty = totDuty + iValue+iEValue+iVValue;
		}


		iSday  = parseInt(iWorkday,10)- parseInt(iWeekday,10)+1 ;
		iEday = iSday+6;
		if (iSday< 0 ){
			iSday=0;
		}
		if(iEday > document.frmPay.hidEday.value){
			iEday=document.frmPay.hidEday.value;
		}
 
		var totDutyH = parseInt(totDuty/60,10); 
		//휴일근무수당
		var WHTime = 0;
		if (totDutyH > 40){
			totDutyH=40;
		}
		if(totDutyH>=15){WHTime= (8*(totDutyH/40))*60;}

		if (totDutyH>=15 && iWHD >= 0 && chkWHD ==0){   
			eval("document.frmPay.iwhWT"+iWHD).value = jsTimeForm(WHTime);  
		}


	}
}

//근무시간 월간 총합계
function jsSetTotTimeSum(){
	var arrValue1, arrValue2, arrValue3, arrValue4, arrValue5;
	var iwt =0;
	var iewt =0;
	var inwt =0;
	var ihwt =0;
	var iwhwt =0;

	for(i=0;i<= document.frmPay.hidEday.value;i++){
		arrValue1 = eval("document.frmPay.iWT"+i).value.split(":");
		arrValue2 = eval("document.frmPay.ieWT"+i).value.split(":");
		arrValue3 = eval("document.frmPay.inWT"+i).value.split(":");
		arrValue4 = eval("document.frmPay.ihWT"+i).value.split(":");
		arrValue5 = eval("document.frmPay.iwhWT"+i).value.split(":");
		iwt = iwt + parseInt(arrValue1[0],10)*60+parseInt(arrValue1[1],10);
		iewt = iewt + parseInt(arrValue2[0],10)*60+parseInt(arrValue2[1],10);
		inwt = inwt + parseInt(arrValue3[0],10)*60+parseInt(arrValue3[1],10);
		ihwt = ihwt + parseInt(arrValue4[0],10)*60+parseInt(arrValue4[1],10);
		iwhwt = iwhwt + parseInt(arrValue5[0],10)*60+parseInt(arrValue5[1],10);
	}


	if(typeof(document.frmPay.iwhWT32) !="undefined"){
		if(document.frmPay.iwhWT32.value !="0"){
			jsAddWH();
			arrValue5 =  document.frmPay.iwhWT32.value.split(":");
			iwhwt=iwhwt+ parseInt(arrValue5[0],10)*60+parseInt(arrValue5[1],10);
		}
	}

	document.frmPay.totWT.value = jsTimeForm(iwt);
	document.frmPay.toteWT.value = jsTimeForm(iewt);
	document.frmPay.totnWT.value = jsTimeForm(inwt);
	document.frmPay.tothWT.value = jsTimeForm(ihwt);
	document.frmPay.totwhWT.value = jsTimeForm(iwhwt);
}
