<script language="JScript" src="/lib/util/editor.js"></script>
<script language="javascript">
<!--
	window.onload = function(){
	//로딩시 포커스 위치 입니다.
		editor.focus();
	//에디터 초기화및 포인터 생성
		ready_edit();
	//내용의 큰따옴표문제로 다른데서 불러다 붙이기
		editor.document.body.innerHTML=content_data.innerHTML;
	}
//-->
</script>
<table border="0" cellpadding="0" class="table_form" width="100%">
<tr>
	<td width="100%">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" height="26" class="table_form">
		<tr>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td>
				<select class="808080_input" onchange="cmdExec('fontname',this[this.selectedIndex].value);" name="font">
					<option value="굴림체" class="heading">굴림체</option>
					<option value="돋움체">돋움체</option>
					<option value="바탕체">바탕체</option>
					<option value="궁서체">궁서체</option>
					<option value="geneva,arial,sans-serif">Arial</option>
					<option value="tahoma">Tahoma</option>
					<option value="courier, monospace">Courier</option>
					<option value="Comic Sans MS">Comic Sans MS</option>
				</select>
			</td>
			<td>
				<select class="808080_input" onchange="cmdExec('fontsize',this[this.selectedIndex].value);" name="size">
					<option selected>Size</option>
					<option value="1">1(8pt)</option>
					<option value="2">2(10pt)</option>
					<option value="3">3(12pt)</option>
					<option value="4">4(14pt)</option>
					<option value="5">5(18pt)</option>
					<option value="6">6(24pt)</option>
					<option value="7">7(36pt)</option>
				</select>
			</td>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_numberlist.gif" id="InsertOrderedList" onClick="cmdExec('InsertOrderedList')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="번호매기기" width="23" height="22"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_balllist.gif" id="InsertUnOrderedList" onClick="cmdExec('InsertUnOrderedList')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="글머리기호" width="23" height="22"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_outdent.gif" id="Outdent" onClick="cmdExec('Outdent')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="내어쓰기" width="23" height="22"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_indent.gif" id="Indent" onClick="cmdExec('Indent')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="들여쓰기" width="23" height="22"></td>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_left.gif" id="JustifyLeft" onClick="cmdExec('JustifyLeft')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="왼쪽정렬" width="23" height="22"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_center.gif" id="JustifyCenter" onClick="cmdExec('JustifyCenter')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="가운데정렬" width="23" height="22"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_right.gif" id="JustifyRight" onClick="cmdExec('JustifyRight')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="오른쪽정렬" width="23" height="22"></td>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_hr.gif" id="hr" onClick="cmdHrInput()" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="수평선 넣기" width="23" height="22"></td>
		</tr>
		</table>
		<table  height="26" cellspacing="0" cellpadding="0" border="0" class="table_form">
		<tr>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_cut.gif" id="Cut" onClick="cmdExec('Cut')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="잘라내기(Ctrl + x)" width="23" height="22"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_copy.gif" id="Copy" onClick="cmdExec('Copy')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="복사(Ctrl + c)" width="23" height="22"></td>
			<td align="center" width="25" class="hand"><img border="0" src="/images/editor/icon_paste.gif" id="Paste" onClick="cmdExec('Paste')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="붙여넣기(Ctrl + v)" width="23" height="22"></td>
			<td class="hand" width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td class="hand" width="25" align="center"><img border="0" src="/images/editor/icon_fontcolor.gif" id="FontColor" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop, 'ForeColor')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="글자색" width="23" height="22"></td>
			<td width="25" align="center"><img border="0" src="/images/editor/icon_backcolor.gif" id="FontColor0" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop, 'BackColor')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="글자배경색" width="23" height="22"></td>
			<td class="hand" width="25" align="center"><img border="0" src="/images/editor/icon_b.gif" id="bold" onClick="cmdExec('bold')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="굵게" width="23" height="22"></td>
			<td class="hand" width="25" align="center"><img border="0" src="/images/editor/icon_i.gif" id="italic" onClick="cmdExec('italic')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="기울임꼴" width="23" height="22"></td>
			<td class="hand" width="25" align="center"><img border="0" src="/images/editor/icon_u.gif" id="underline" onClick="cmdExec('underline')" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" class="img_act" alt="밑줄" width="23" height="22"></td>
			<td width="10" align="center"><img border="0" src="/images/editor/icon_bar.gif" width="2" height="26"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_link.gif" id="Link" onClick="ShowLinkBox(event.clientX, event.clientY+document.body.scrollTop)" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="하이퍼링크 삽입" width="23" height="22"></td>
			<td width="25" align="center" class="hand"><img border="0" src="/images/editor/icon_image.gif" id="Image" onClick="ShowImgLinkBox(event.clientX, event.clientY+document.body.scrollTop)" chk="0" OnMouseOver="button_over(this)" OnMouseOut="button_out(this)" onmousedown="button_down(this);" alt="그림 삽입" width="23" height="22"></td>
		</tr>
		</table>
	</td>
</tr>
<tr id="sector_1" style="display:block" chk="0">
	<td width="100%">
		<iframe name="editor" id="editor" OnFocus="divLayerOFF();" OnBlur="NowSpace.SaveSelection();" height="<%=editor_height%>" width="<%=editor_width%>" marginwidth="5" marginheight="5" border="1" frameborder="1" chk="0"></iframe>
	</td>
</tr>
<tr id="sector_2" style="display:none" chk="0">
	<td width="100%">
		<iframe name="preview" id="preview" height="<%=editor_height%>" width="<%=editor_width%>" marginwidth="5" marginheight="5" border="1" frameborder="1" chk="0"></iframe>
</td>
</tr>
<tr>
	<td width="100%">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="19">
		<tr>
			<td width="42" align="center">
				<div id="edit_default" class="bt_btn_click" OnClick="select_btn(this,'click', 1)" OnMouseOver="select_btn(this,'over', 1)" OnMouseOut="select_btn(this,'out', 1)" chk="1">
				<img border="0" src="/images/editor/icon_edit.gif" width="40" height="17" class='HandCursor'></div>
			</td>
			<td width="42" align="center">
				<div id="edit_html" class="bt_btn_out" OnClick="select_btn(this,'click', 2)" OnMouseOver="select_btn(this,'over', 2)" OnMouseOut="select_btn(this,'out', 2)" chk="0">
				<img border="0" src="/images/editor/icon_html.gif" width="50" height="17" class='HandCursor'></div>
			</td>
			<td width="70" align="center">
				<div id="edit_preview" class="bt_btn_out" chk="0"></div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr style="display:none">
	<td width="100%">
<input type="hidden" name="brd_content" value="">
<input type="hidden" name="chk_editor" value="editor">
<div name='content_data' id='content_data' style='position:absolute;visibility:hidden;left:0;top:0;'><%=brd_content%></div>	
	</td>
</tr>
</table>