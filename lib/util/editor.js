<!--
	//<p></p> 테그시 <br>처럼 보이게 하기위해..
	//var innerHeader = "<style>P {margin-top:2px;margin-bottom:2px;} \nbody {font-size:9pt; font-family:굴림;}</style>\n";
	var innerHeader = "<style>body {font-size:9pt; font-family:굴림;}</style>\n";

	function ready_edit(){
		//선택 영역 공간들 저장 부분
		//*** 임시적으로 사용 안함 ***
		NowSpace = new Space();

		editor.document.designMode="On"

		editor.document.open();
		editor.document.write(innerHeader);
		editor.document.close();
	}

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	해당 커멘트 입력 부분..
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function cmdExec(cmd, opt){
		NowSpace.RestoreSelection();
		
		var tt = editor.document.selection.createRange();
				
		if (sector_1.chk == '1'){
			alert('HTML, 미리보기 상태에서는 수정을 하실수가 없습니다.');
		}else{
			NowSpace.RestoreSelection();
		
			if (opt==null){
				editor.document.execCommand(cmd);
			}else{
				editor.document.execCommand(cmd,"",opt);
			}
		}
	}

	function changeColor(cmd, opt){
		ColorBox.style.visibility = 'hidden';
		cmdExec(cmd, opt);
	}
	
	function cmdHrInput(){
		if (sector_1.chk == '1'){
			alert('HTML, 미리보기 상태에서는 수정을 하실수가 없습니다.');
		}else{
		
			//--------------------------------------
			//	포커스가 입력창에 없을 때 강제적으로 포커스 이동
			//--------------------------------------
				NowSpace.RestoreSelection();
		
				if(NowSpace.selection){
					var aa = NowSpace.selection.parentElement();
					if(aa.style.topmargin != "12px"){
						editor.focus();
					}
				}
			//----------------------------------------	
	
			var EditorCtrl = editor.document.selection.createRange();
				EditorCtrl.pasteHTML("<hr>");
		}
	}
	
	function cmdImageLink(){
	
		var linkURL = new Image;
			linkURL.src = document.all.img_url.value;
		
		if (linkURL.src=='http:///' || linkURL.src=='' || linkURL.src==null || linkURL.width==0 || linkURL.height==0){
			
			var answer = confirm('입력하시려는 그림이 해당 링크 주소에 존재하지 않습니다.\n\n 다시 입력 하시겠습니까?.');
			
			if(answer){
				document.all.img_url.select();
			}else{
				divLayerOFF();
			}
			
		}else{
			divLayerOFF();

			//--------------------------------------
			//	포커스가 입력창에 없을 때 강제적으로 포커스 이동
			//--------------------------------------

			NowSpace.RestoreSelection();
		
			if(NowSpace.selection){
				var aa = NowSpace.selection.parentElement();
				if(aa.style.topmargin != "12px"){
					editor.focus();
				}
			}
			//----------------------------------------	
			var EditorCtrl = editor.document.selection.createRange();
			EditorCtrl.pasteHTML("<IMG src='"+linkURL.src+"'>");
		
			//cmdExec("InsertImage",linkURL.src);
		}		
	}

	function cmdLink(){
	
		var linkURL = document.all.url_str.value;
		
			divLayerOFF();

			//--------------------------------------
			//	포커스가 입력창에 없을 때 강제적으로 포커스 이동
			//--------------------------------------

			NowSpace.RestoreSelection();
		
			if(NowSpace.selection){
				var aa = NowSpace.selection.parentElement();
				if(aa.style.topmargin != "12px"){
					editor.focus();
				}
			}
			//----------------------------------------	
			var EditorCtrl = editor.document.selection.createRange();
				
				if (EditorCtrl.htmlText==''){
					EditorCtrl.pasteHTML("<A HREF='"+linkURL+"' TARGET='_blank'>"+ linkURL +"</A>");
				}else{
					EditorCtrl.pasteHTML("<A HREF='"+linkURL+"' TARGET='_blank'>"+ EditorCtrl.htmlText +"</A>");
				}
	}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	글자, 그림 등 선택 영역 저장 부분
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	function Space(){
		this.selection    = null;
		this.selection2    = null;
		this.RestoreSelection = Space_RestoreSelection;
		this.SaveSelection  = Space_SaveSelection;
		this.GetSelection  = Space_GetSelection;
	}

	function Space_RestoreSelection() {
		if (this.selection) {
			this.selection.select();
		}
	}

	function Space_GetSelection() {
		var oSelected = this.selection;
		if (!oSelected) {
			oSelelected = editor.document.selection.createRange();
			oSelelected.type = editor.document.selection.type;
		}
		return oSelected;
	}

	function Space_SaveSelection() {
		NowSpace.selection = editor.document.selection.createRange();
		NowSpace.selection.type = editor.document.selection.type;
	}

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	글자색, 배경색 선택
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function ShowColorBox(posX, posY, cmd){
		
		if (sector_1.chk == '1'){
			alert('HTML, 미리보기 상태에서는 수정을 하실수가 없습니다.');
		}else{
			if (ColorBox.style.visibility != 'visible'){
				ColorBox.innerHTML = tagColor(cmd);
				ColorBox.style.visibility = 'visible';
				ColorBox.style.left = (posX - 12);
				ColorBox.style.top = (posY + 12);
			}else{
				divLayerOFF();
			}
		}
	}
	
	function tagColor(cmd){
		// cmd == "forecolor" -> 폰트 색 
		// cmd == "backcolor" -> 배경 색
		
		//화면에 보이는 색상톤.. 
		var colortone = new Array(15);
			colortone[0] = new Array('#000000','#ffffff','#008000','#800000','#ac8295','#808000','#000080','#800080','#808080','#c0c0c0');
			colortone[1] = new Array('#ffff00','#00ff00','#00ffff','#ff00ff','#ff0000','#0000ff','#008080','#ed8602','#0099ff','#9900ff');
			colortone[2] = new Array('#ffffff','#e5e4e4','#d9d8d8','#c0bdbd','#a7a4a4','#8e8a8b','#827e7f','#767173','#5c585a','#000000');
			colortone[3] = new Array('#fefcdf','#fef4c4','#feed9b','#fee573','#ffed43','#f6cc0b','#e0b800','#c9a601','#ad8e00','#8c7301');
			colortone[4] = new Array('#ffded3','#ffc4b0','#ff9d7d','#ff7a4e','#ff6600','#e95d00','#d15502','#ba4b01','#a44201','#8d3901');
			colortone[5] = new Array('#ffd2d0','#ffbab7','#fe9a95','#ff7a73','#ff483f','#fe2419','#f10b00','#d40a00','#940000','#6d201b');
			colortone[6] = new Array('#ffdaed','#ffb7dc','#ffa1d1','#ff84c3','#ff57ac','#fd1289','#ec0078','#d6006d','#bb005f','#9b014f');
			colortone[7] = new Array('#fcd6fe','#fbbcff','#f9a1fe','#f784fe','#f564fe','#f546ff','#f328ff','#d801e5','#c001cb','#8f0197');
			colortone[8] = new Array('#e2f0fe','#c7e2fe','#add5fe','#92c7fe','#6eb5ff','#48a2ff','#2690fe','#0162f4','#013add','#0021b0');
			colortone[9] = new Array('#d3fdff','#acfafd','#7cfaff','#4af7fe','#1de6fe','#01deff','#00cdec','#01b6de','#00a0c2','#0084a0');
			colortone[10] = new Array('#edffcf','#dffeaa','#d1fd88','#befa5a','#a8f32a','#8fd80a','#79c101','#3fa701','#307f00','#156200');
			colortone[11] = new Array('#d4c89f','#daad88','#c49578','#c2877e','#ac8295','#c0a5c4','#969ac2','#92b7d7','#80adaf','#9ca53b');

		var strHTML = "";
			strHTML = strHTML + "<table cellpadding=2 cellspacing=0 border=1 style='border-collapse: collapse' bgcolor='#FFFFFF'><tr><td><table cellpadding=0 cellspacing=0 border=0>";

			for (var i=0; i<11; i++){
				strHTML = strHTML + "<tr>";
				
				for(var j=0; j<10; j++){
					strHTML = strHTML + "<td onmouseover=this.style.backgroundColor='blue' onmouseout=this.style.backgroundColor='' class='hand' title='" + colortone[i][j] + "'><table cellpadding=0 cellspacing=1 border=0><tr><td bgcolor='" + colortone[i][j] + "' onclick='changeColor(\"" + cmd + "\", \"" + colortone[i][j] + "\");' width=10 height=10></td></tr></table></td>";
				}   
				strHTML = strHTML + "</tr>";
			}
			strHTML = strHTML + "</table></td></tr></table>";
			
			return strHTML;
	}
	
	function divLayerOFF(){
	
		try{
			ColorBox.style.visibility = 'hidden';
			imgLinkBox.style.visibility = 'hidden';
			LinkBox.style.visibility = 'hidden';
		}catch(e){}
	}
	
		
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	이미지 링크 선택
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function ShowImgLinkBox(posX, posY,cmd){
		
		if (sector_1.chk == '1'){
			alert('HTML 상태에서는 수정을 하실수가 없습니다.');
		}else{
			if (imgLinkBox.style.visibility != 'visible'){
				imgLinkBox.innerHTML = imgLinkInput();
				imgLinkBox.style.visibility = 'visible';
				imgLinkBox.style.left = (posX - 12);
				imgLinkBox.style.top = (posY + 12);
			
				document.all.img_url.focus();
				document.all.img_url.select();
			}else{
				divLayerOFF();
			}
		}
	}
	
	function imgLinkInput(cmd){
		
		if (cmd=='' || cmd=='undefined' || cmd==null){
			cmd='http://';
		}
		
		var strHTML = "";
			strHTML = strHTML + "<table border='1' cellspacing='0' width='300' style='border-collapse: collapse' bordercolor='#0A246A' cellpadding='5'><tr><td width='100%' bgcolor='#ffffff'><span style='font-size: 9pt'>인터넷 상에 올려진 이미지만 가능 합니다.<br>";
			strHTML = strHTML + "삽입할 이미지의 인터넷주소(URL)을 넣어 주세요.<br><font color='red'>『<b>http://</b>』를 꼭쓰세요.<br><input type='text' name='img_url' size='26' class='808080_input' value='"+ cmd +"'>&nbsp; <img src='/images/editor/icon_ok.gif' class='hand' OnClick='cmdImageLink();'> <img src='/images/editor/icon_cancel.gif' class='hand' OnClick='divLayerOFF();'></span></td></tr></table>";
			
		return strHTML;
	}


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	이미지 선택/업로드 박스 오픈
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function OpenImgSelBox(){
		
		if (sector_1.chk == '1'){
			alert('HTML 상태에서는 수정을 하실수가 없습니다.');
		}else{
			window.open("pop_imgbox/imgSelBox.asp","ImgSelect","width=420,height=400,scrollbars=yes");
		}
	}

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	텍스트 링크 선택
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function ShowLinkBox(posX, posY,cmd){
		
		if (sector_1.chk == '1'){
			alert('HTML, 미리보기 상태에서는 수정을 하실수가 없습니다.');
		}else{
			if (LinkBox.style.visibility != 'visible'){
				LinkBox.innerHTML = LinkInput(cmd);
				LinkBox.style.visibility = 'visible';
				LinkBox.style.left = (posX - 12);
				LinkBox.style.top = (posY + 12);
			
				document.all.url_str.focus();
				document.all.url_str.select();
			}else{
				divLayerOFF();
			}
		}
	}
	
	function LinkInput(cmd){
		
		if (cmd=='' || cmd=='undefined' || cmd==null){
			cmd='http://';
		}
		
		var strHTML = "";
			strHTML = strHTML + "<table border='1' cellspacing='0' width='300' style='border-collapse: collapse' bordercolor='#0A246A' cellpadding='5'><tr><td width='100%' bgcolor='#ffffff'><span style='font-size: 9pt'>링크하실 페이지 주소를 적어 주세요.<br>";
			strHTML = strHTML + "링크할할 인터넷주소(URL)을 넣어 주세요.<br><font color='red'>『<b>http://</b>』를 꼭쓰세요.<br><input type='text' name='url_str' size='26' class='808080_input' value='"+ cmd +"'>&nbsp; <img src='/images/editor/icon_ok.gif' class='hand' OnClick='cmdLink();'> <img src='/images/editor/icon_cancel.gif' class='hand' OnClick='cmdExec(&quot;unlink&quot;);divLayerOFF();'></span></td></tr></table>";
			
		return strHTML;
	}


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	버튼 변화
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	function button_over(btn_obj){
			btn_obj.style.borderBottom = "buttonshadow solid 1px";
			btn_obj.style.borderLeft = "buttonhighlight solid 1px";
			btn_obj.style.borderRight = "buttonshadow solid 1px";
			btn_obj.style.borderTop = "buttonhighlight solid 1px";
	}

	function button_out(btn_obj){
			btn_obj.style.borderColor = "threedface";
			btn_obj.style.border="0px";
	}

	function button_down(btn_obj){
			btn_obj.style.borderBottom = "buttonhighlight solid 1px";
			btn_obj.style.borderLeft = "buttonshadow solid 1px";
			btn_obj.style.borderRight = "buttonhighlight solid 1px";
			btn_obj.style.borderTop = "buttonshadow solid 1px";
	}
	
	function button_up(btn_obj){
			btn_obj.style.borderBottom = "buttonshadow solid 1px";
			btn_obj.style.borderLeft = "buttonhighlight solid 1px";
			btn_obj.style.borderRight = "buttonshadow solid 1px";
			btn_obj.style.borderTop = "buttonhighlight solid 1px";
			btn_obj = null; 
	}
	
	
	function select_btn(obj, order, sector){
		
		var cname;
		var chk = obj.chk;
		
		if (order=='over'){
			if (chk=='0'){
				cname='bt_btn_over';
			}else{
				cname='';
			}
		}else if(order=='out'){
			if (chk=='0'){
				cname='bt_btn_out';
			}else{
				cname='';
			}
		}else if(order=='click'){
			edit_default.className = 'bt_btn_out';
			edit_default.chk = '0';
			edit_html.className = 'bt_btn_out';
			edit_html.chk = '0';
			edit_preview.className = 'bt_btn_out';
			edit_preview.chk = '0';
			
			obj.chk=1;
			cname='bt_btn_click';
			
			if(sector==1){
					sector_1.style.display = 'block';
					sector_2.style.display = 'none';
				
				if (sector_1.chk != 0){
					sector_1.chk = 0;
					sector_2.chk = 0;
					editor.document.body.innerHTML = editor.document.body.innerText;
				}
				
				editor.focus();
			}else if(sector==2){
					sector_1.style.display = 'block';
					sector_2.style.display = 'none';
				
				if (sector_1.chk != 1){
					editor.document.body.innerText = editor.document.body.innerHTML;
					sector_1.chk = 1;
					sector_2.chk = 0;
				}
				
				editor.focus();
			}else if(sector==3){
				if (sector_2.chk != 1){
					sector_1.style.display = 'none';
					sector_2.style.display = 'block';
				
					if (sector_1.chk==0){
						sector_1.chk = 0;
						preview.document.open();
							preview.document.writeln("\n<html>\n<head>\n"+innerHeader+"</head>\n<body>\n");
							preview.document.writeln(editor.document.body.innerHTML);
						preview.document.close("</body>\n</html>\n");
					}else{
						sector_1.chk = 1;
						preview.document.open();
							preview.document.writeln("\n<html>\n<head>\n"+innerHeader+"</head>\n<body>\n");
							preview.document.writeln(editor.document.body.innerText);
							preview.document.writeln("</body>\n</html>\n");
						preview.document.close();
					}
	
					sector_2.chk = 1;
				}
				preview.focus();
			}
		}
				
		if (cname!=''){
			obj.className = cname;
		}
	}	
	
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	폼 전송 관련..
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function form_reset(){
		editor.document.body.innerHTML = '';
	}


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

document.write("<div id='ColorBox' style='position:absolute;visibility:hidden;left:200;top:100;'></div>");
document.write("<div id='imgLinkBox' style='position:absolute;visibility:hidden;left:200;top:100;'></div>");
document.write("<div id='LinkBox' style='position:absolute;visibility:hidden;left:200;top:100;'></div>");

//-->