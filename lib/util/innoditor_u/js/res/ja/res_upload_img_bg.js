/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// res_upload_img_bg.js
//						
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

var res_item = new Array();
var res_warning_msg = new Array();
var res_repeat_text = new Array();

var res_filename = "選択したファイル名";


res_item[0] = "JPG,GIF,PNGファイル<br>(1024K以下のみ可能)";
res_item[1] = "<b>背景イメージの繰り返し(選択)</b>";
res_item[2] = "設定された背景イメージを除去します";


res_warning_msg[0] = "ファイル選択で, 先にイメージファイルを選んでください";
res_warning_msg[1] = "イメージファイル(jpg,gif,png)のみアップできます";
res_warning_msg[2] = "アップロードに失敗しました\nイメージファイルではないか, 用量オーバーです";
res_warning_msg[3] = "先にイメージをアップロードしてください";

res_repeat_text[0] = "繰り返し(基本値)";
res_repeat_text[1] = "繰り返しなし";
res_repeat_text[2] = "横繰り返し";
res_repeat_text[3] = "縦繰り返し";