/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// browser.js -	브라우저 종류 검사 및 버전 검사
//						
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function cm_bwcheck()
{
	this.agent = navigator.userAgent.toLowerCase();

	this.ie13 = (this.agent.indexOf("msie 13") > -1);
	this.ie12 = (this.agent.indexOf("msie 12") > -1 && !this.ie13);
	this.ie11 = (this.agent.indexOf("msie 11") > -1 && !this.ie12 && !this.ie13);
	this.ie10 = (this.agent.indexOf("msie 10") > -1 && !this.ie11 && !this.ie12 && !this.ie13);
	this.ie9 = (this.agent.indexOf("msie 9") > -1 && !this.ie10 && !this.ie11 && !this.ie12 && !this.ie13);
	this.ie8 = (this.agent.indexOf("msie 8") > -1 && !this.ie9 && !this.ie10 && !this.ie11 && !this.ie12 && !this.ie13);
	this.ie7 = (this.agent.indexOf("msie 7") > -1 && !this.ie8 && !this.ie9 && !this.ie10 && !this.ie11 && !this.ie12 && !this.ie13);
	this.ie6 = (this.agent.indexOf("msie 6") > -1 && !this.ie7 && !this.ie8 && !this.ie9 && !this.ie10 && !this.ie11 && !this.ie12 && !this.ie13);
	this.ie = (this.ie6 || this.ie7 || this.ie8 || this.ie9 || this.ie10 || this.ie11 || this.ie12 || this.ie13);

	this.ff30 = (this.agent.indexOf("firefox/30") > -1);
	this.ff29 = (this.agent.indexOf("firefox/29") > -1);
	this.ff28 = (this.agent.indexOf("firefox/28") > -1);
	this.ff27 = (this.agent.indexOf("firefox/27") > -1);
	this.ff26 = (this.agent.indexOf("firefox/26") > -1);
	this.ff25 = (this.agent.indexOf("firefox/25") > -1);
	this.ff24 = (this.agent.indexOf("firefox/24") > -1);
	this.ff23 = (this.agent.indexOf("firefox/23") > -1);
	this.ff22 = (this.agent.indexOf("firefox/22") > -1);
	this.ff21 = (this.agent.indexOf("firefox/21") > -1);
	this.ff20 = (this.agent.indexOf("firefox/20") > -1);
	this.ff19 = (this.agent.indexOf("firefox/19") > -1);
	this.ff18 = (this.agent.indexOf("firefox/18") > -1);
	this.ff17 = (this.agent.indexOf("firefox/17") > -1);
	this.ff16 = (this.agent.indexOf("firefox/16") > -1);
	this.ff15 = (this.agent.indexOf("firefox/15") > -1);
	this.ff14 = (this.agent.indexOf("firefox/14") > -1);
	this.ff13 = (this.agent.indexOf("firefox/13") > -1);
	this.ff12 = (this.agent.indexOf("firefox/12") > -1);
	this.ff11 = (this.agent.indexOf("firefox/11") > -1);
	this.ff10 = (this.agent.indexOf("firefox/10") > -1);
	this.ff9 = (this.agent.indexOf("firefox/9") > -1);
	this.ff8 = (this.agent.indexOf("firefox/8") > -1);
	this.ff7 = (this.agent.indexOf("firefox/7") > -1);
	this.ff6 = (this.agent.indexOf("firefox/6") > -1);
	this.ff5 = (this.agent.indexOf("firefox/5") > -1);
	this.ff4 = (this.agent.indexOf("firefox/4") > -1);
	this.ff3 = (this.agent.indexOf("firefox/3") > -1);
	this.ff2 = (this.agent.indexOf("firefox/2") > -1);
	this.ff = (this.ff2 || this.ff3 || this.ff4 || this.ff5 || this.ff6 || this.ff7 || this.ff8 || this.ff9 || this.ff10 
				|| this.ff11 || this.ff12 || this.ff13 || this.ff14 || this.ff15 || this.ff16 || this.ff17 || this.ff18 || this.ff19 || this.ff20 
				|| this.ff21 || this.ff22 || this.ff23 || this.ff24 || this.ff25 || this.ff26 || this.ff27 || this.ff28 || this.ff29 || this.ff30);

	this.wk = (this.agent.indexOf("webkit") > -1);
	this.cr = false;
	if(this.wk)
	{
		this.cr = (this.agent.indexOf("chrome") > -1);
	}

	this.op = (this.agent.indexOf("opera") > -1);


	if(this.ie)
	{
		if(this.ie6) { this.verInfo = 6; }
		else if(this.ie7) { this.verInfo = 7; }
		else if(this.ie8) { this.verInfo = 8; }
		else if(this.ie9) { this.verInfo = 9; }
		else if(this.ie10) { this.verInfo = 10; }
		else if(this.ie11) { this.verInfo = 11; }
		else if(this.ie12) { this.verInfo = 12; }
		else if(this.ie13) { this.verInfo = 13; }
	}

	if(this.ff)
	{
		if(this.ff2) { this.verInfo = 2; }
		else if(this.ff3) { this.verInfo = 3; }
		else if(this.ff4) { this.verInfo = 4; }
		else if(this.ff5) { this.verInfo = 5; }
		else if(this.ff6) { this.verInfo = 6; }
		else if(this.ff7) { this.verInfo = 7; }
		else if(this.ff8) { this.verInfo = 8; }
		else if(this.ff9) { this.verInfo = 9; }
		else if(this.ff10) { this.verInfo = 10; }
		else if(this.ff11) { this.verInfo = 11; }
		else if(this.ff12) { this.verInfo = 12; }
		else if(this.ff13) { this.verInfo = 13; }
		else if(this.ff14) { this.verInfo = 14; }
		else if(this.ff15) { this.verInfo = 15; }
		else if(this.ff16) { this.verInfo = 16; }
		else if(this.ff17) { this.verInfo = 17; }
		else if(this.ff18) { this.verInfo = 18; }
		else if(this.ff19) { this.verInfo = 19; }
		else if(this.ff20) { this.verInfo = 20; }
		else if(this.ff21) { this.verInfo = 21; }
		else if(this.ff22) { this.verInfo = 22; }
		else if(this.ff23) { this.verInfo = 23; }
		else if(this.ff24) { this.verInfo = 24; }
		else if(this.ff25) { this.verInfo = 25; }
		else if(this.ff26) { this.verInfo = 26; }
		else if(this.ff27) { this.verInfo = 27; }
		else if(this.ff28) { this.verInfo = 28; }
		else if(this.ff29) { this.verInfo = 29; }
		else if(this.ff30) { this.verInfo = 30; }
	}

	this.bw = (this.ie || this.ff || this.wk || this.op);


	if(navigator.userLanguage)
	{
		this.language = navigator.userLanguage.toLowerCase();
	}
	else if (navigator.language)
	{
		this.language = navigator.language.toLowerCase();
	}
	else
	{
		this.language = null;
	}

	return this;
}

var g_browserCHK = new cm_bwcheck();
