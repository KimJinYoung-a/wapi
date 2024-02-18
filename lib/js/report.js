function SelReport(comp,targetfrm){
	
	if (comp.value=='year'){
		targetfrm.yyyy1.disabled = false;
		targetfrm.yyyy2.disabled = false;
		targetfrm.mm1.disabled = true;
		targetfrm.mm2.disabled = true;
		targetfrm.dd1.disabled = true;
		targetfrm.dd2.disabled = true;
	}else if (comp.value=='month'){
		targetfrm.yyyy1.disabled = false;
		targetfrm.yyyy2.disabled = false;
		targetfrm.mm1.disabled = false;
		targetfrm.mm2.disabled = false;
		targetfrm.dd1.disabled = true;
		targetfrm.dd2.disabled = true;
	}else if (comp.value=='day'){
		targetfrm.yyyy1.disabled = false;
		targetfrm.yyyy2.disabled = false;
		targetfrm.mm1.disabled = false;
		targetfrm.mm2.disabled = false;
		targetfrm.dd1.disabled = false;
		targetfrm.dd2.disabled = false;
	}else if (comp.value=='week'){
		targetfrm.yyyy1.disabled = false;
		targetfrm.yyyy2.disabled = false;
		targetfrm.mm1.disabled = false;
		targetfrm.mm2.disabled = false;
		targetfrm.dd1.disabled = false;
		targetfrm.dd2.disabled = false;
	}else if (comp.value=='time'){
		targetfrm.yyyy1.disabled = false;
		targetfrm.yyyy2.disabled = true;
		targetfrm.mm1.disabled = false;
		targetfrm.mm2.disabled = true;
		targetfrm.dd1.disabled = false;
		targetfrm.dd2.disabled = true;
	}
}

function drawMonthReport(val){
	var v;
	v =	"<OBJECT classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 WIDTH=580 HEIGHT=220>";
	v =	v +	"<PARAM NAME=movie VALUE='/graph/get_graph_month.swf?" + val + "'>"; //pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996'>";
	v =	v +	"<PARAM NAME=loop VALUE=false>";
	v =	v +	"<PARAM NAME=menu VALUE=false>";
	v =	v +	"<PARAM NAME=quality VALUE=high>";
	v =	v +	"<PARAM NAME=bgcolor VALUE=#FFFFFF>";
	v =	v +	"<PARAM NAME=wmode VALUE=Transparent>";
	v =	v +	"</Object>";
	//v =	v +	"<EMBED src='/graph/get_graph_month.swf?pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996' loop=false menu=false quality=high bgcolor=#FFFFFF wmode=Transparent  WIDTH=580 HEIGHT=220 TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash'></EMBED>";
	document.write(v);
}

function drawDayReport(val){
	var v;
	v =	"<OBJECT classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 WIDTH=580 HEIGHT=220>";
	v =	v +	"<PARAM NAME=movie VALUE='/graph/get_graph_month.swf?pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996'>";
	v =	v +	"<PARAM NAME=loop VALUE=false>";
	v =	v +	"<PARAM NAME=menu VALUE=false>";
	v =	v +	"<PARAM NAME=quality VALUE=high>";
	v =	v +	"<PARAM NAME=bgcolor VALUE=#FFFFFF>";
	v =	v +	"<PARAM NAME=wmode VALUE=Transparent>";
	v =	v +	"</Object>";
	//v =	v +	"<EMBED src='/graph/get_graph_month.swf?pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996' loop=false menu=false quality=high bgcolor=#FFFFFF wmode=Transparent  WIDTH=580 HEIGHT=220 TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash'></EMBED>";
	document.write(v);
}