/////////////////////////////////////////////////////////////////////////////////////////////////////
// editor 갯수(1개~20개) 및 editor를 로딩할 영역의 ID 설정
// 페이지별로 editor의 갯수를 설정하고자 하는 경우는 페이지별로 선언할 것
// 사이트 전체에 일괄적으로 1개 또는 N개 로딩하는 경우면 아래부분의 주석을 제거하고 적용

//var g_arrSetEditorArea = new Array();
//g_arrSetEditorArea[0] = "EDITOR_AREA_CONTAINER";// 이노디터를 위치시킬 영역의 ID값 설정
/////////////////////////////////////////////////////////////////////////////////////////////////////



// skin 선택(0~9사이의 skin 선택)
var g_nSkinNumber = 0;

var g_strPath_Image = "/lib/util/innoditor/image/";
var g_strPath_JS = "/lib/util/innoditor/js/";
var g_strPath_CSS = "/lib/util/innoditor/css/";
var g_strPath_Property = "/lib/util/innoditor/";
var g_strPath_License = "/lib/util/innoditor/";

var g_nEditorWidth = 670;
var g_nEditorHeight = 600;


// 제품 도움말 URL
var g_strHelpPageURL = "/lib/util/innoditor/pop_help.html";


// 제품정보 URL
var g_strProductInfoURL = "/lib/util/innoditor/pop_productinfo.html";
//var g_strProductInfoURL = "/lib/util/innoditor/pop_productinfo_en.html";// 영문모드 사용시


// Image 업로드 Page URL
//var g_strUploadImageURL = "/lib/util/innoditor/pop_simple_img.html";//예제
var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.asp"; //2013-01-17 진영 수정
//var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.html";
//var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img_en.html";// 영문모드 사용시
//var g_strUploadImageURL = "/lib/util/innoditor/pop_link_img.html";// 외부링크만 사용시(한글)
//var g_strUploadImageURL = "/lib/util/innoditor/pop_link_img_en.html";// 외부링크만 사용시(영문)


// Flash 업로드 Page URL
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_simple_flash.html";//예제
var g_strUploadFlashURL = "/lib/util/innoditor/pop_upload_flash.html";
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_upload_flash_en.html";// 영문모드 사용시
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_link_flash.html";// 외부링크만 사용시(한글)
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_link_flash_en.html";// 외부링크만 사용시(영문)


// Media 업로드 Page URL
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_simple_media.html";//예제
var g_strUploadMediaURL = "/lib/util/innoditor/pop_upload_media.html";
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_upload_media_en.html";// 영문모드 사용시
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_link_media.html";// 외부링크만 사용시(한글)
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_link_media_en.html";// 외부링크만 사용시(영문)


// 배경 Image 업로드 및 설정 Page URL
//var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_simple_img_bg.html";//예제
var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_upload_img_bg.html";
//var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_upload_img_bg_en.html";// 영문모드 사용시


// 문서 Templete 삽입 Page URL
var g_strInsertDocTempleteURL = "/lib/util/innoditor/pop_doc_templete.html";
//var g_strInsertDocTempleteURL = "/lib/util/innoditor/pop_doc_templete_en.html";// 영문모드 사용시


// 속성 Page URL
var g_strPropertyPageURL = "/lib/util/innoditor/pop_property.html";
//var g_strPropertyPageURL = "/lib/util/innoditor/pop_property_en.html";// 영문모드 사용시


// 미리보기 Page URL
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview.html";
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_en.html";// 영문모드 사용시
var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_x.html";// XHTML 출력방식으로 설정한 경우
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_x_en.html";// XHTML 출력방식으로 설정한 경우(영문)



// 라이센스
var g_arrDomainName = new Array();
g_arrDomainName[0] = "localhost";
g_arrDomainName[1] = "scm.10x10.co.kr";
g_arrDomainName[2] = "testscm.10x10.co.kr";

var g_arrLicenseKey = new Array();
g_arrLicenseKey[0] = "Mv5Oi$BZ+q3Pm/Lq4h@MX4Nh#AYs26EYo&@Tbq3Pm/Lhf+Ap7Y{D";
g_arrLicenseKey[1] = "Ak1Rs9ap!5FUpl<nFwO#!FgKkuP*a:-{U/d>qK.DuM~V.a7:-2Ol.Kh&j4g?pJ%BYr*f*Kk0";
g_arrLicenseKey[2] = "y@b%Km4VxAl1DUhy1MPq7XyAb*p7Y{D4Kl2S#41@3^5hbO&_9oJxT.c<pJ}W**QwDjLO&_9oJapJ}W1P[4Yj+If$Eb}A1VT%BYrqE";



// 메뉴바 show 또는 hidden 만 지원(메뉴레이어)
var g_bCustomize_MenuBar_Display = true;

// Bottom Tab바 show 또는 hidden 만 지원(미리보기,편집창,소스창 버튼)
var g_bCustomize_TabBar_Display = true;

// 첫번째 툴바 show 또는 hidden 설정(세부 항목은 따로 설정)
var g_bCustomize_ToolBar1_Display = true;

// 두번째 툴바 show 또는 hidden 설정(세부 항목은 따로 설정)
var g_bCustomize_ToolBar2_Display = true;

// 세번째 툴바 show 또는 hidden 설정(세부 항목은 따로 설정)
var g_bCustomize_ToolBar3_Display = true;

// 사용자 정의 툴바(이노디터에서 제공되는 기능을 사용하지 않을 경우 이노디터 Interface 만 연동)
var g_bCustomize_CustomToolbar_Display = false;
var g_bCustomize_CustomToolbar_Layout = 0;// 0 - 해당사항 없음, 1 - Top(툴바 상단 영역), 2 - Bottom(툴바 하단 영역)
var g_bCustomize_CustomToolbar_HTML = "";// 사용자정의 툴바에 들어갈 HTML 내용(<table> ~ </table> : table로 시작하여 table로 끝나야 함)


// 툴바 셋팅용 변수 선언 (툴바 셋팅 세부 항목 샘플은 customize_ui.js 참조)
var g_arrCustomToolbar1 = new Array();
var g_arrCustomToolbar2 = new Array();
var g_arrCustomToolbar3 = new Array();


document.write('<script type="text/javascript" src="' + g_strPath_JS + 'browser.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr489343715.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr670454868.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr873475877.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr528318566.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr696495397.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr988789177.js"></scrip' +'t>');
