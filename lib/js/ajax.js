// AJAX 프로그램
//
// 설명
// 웹페이지에서 비동기적으로 사용자 페이지에 정보를 표시할때 사용한다.
// 아래에 사용법을 표시한다.
// a.asp 는 사용자페이지이며, b.asp 는 사용자페이지에 표시할 데이타를 xml 로 보내주는 서버페이지이다.
//
// startAjaxSample(param1) 는 startAjaxSample(param1,param2, param3) 등의 형태로 변경이 가능하다.
// 기능은 불러올 URL 을 세팅하고( initializeURL() ), 리턴시 실행할 스크립트함수명을 세팅하며( initializeReturnFunction() ), AJAX 를 실행한다.( startRequest() )
//
// 리턴시 실행할 스크립트함수명은 디폴트로 processAjax() 며, initializeReturnFunction() 를 이용하여 변경 가능하다.
// 기능은 xmlDoc 에 세팅된 데이타를 사용자화면에 표시한다.
//
// 추가 : 200ms 의 딜레이 추가(select 테그 또는 더블클릭등에 대처)
// 추가 : 특수문자 처리(데이타를 받아오는 페이지(아래 샘플에서는 b.asp)에서 <,>,& 에 대해, &lt;,&gt;,&amp; 등으로 변환하여 전송한다.)
// 추가 : 리턴시 실행할 스크립트함수명을 변경할수 있다.(한페이지에 여러개의 ajax 가 사용될수 있다)
// 추가 : 에러코드를 확인할수 있는 방법을 제공한다. 디폴트로 alert() 를 실행한다.
// 추가 : 딜레이를 100ms 로 줄인다.
//
// 사용법
/*

a.asp
// ============================================================================
<script language="JavaScript" src="/js/ajax.js"></script>
<script>
function startAjaxSample(param1) {
        initializeURL("b.asp?id=" + param1);
        initializeReturnFunction("processAjaxSample()");
        initializeErrorFunction("onErrorAjaxSample()");
        startRequest();
}

function processAjaxSample() {
        // 원본 텍스트
        // xmlHttp.responseText

        var length = xmlDoc.getElementsByTagName("value1").length;
        if (length < 1) {
                alert("데이타가 존재하지 않습니다.");
                return;
        }

        for (var i = 0; i < length; i++) {
                alert(xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue);
                alert(xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue);
        }
}

function onErrorAjaxSample() {
        alert("ERROR : " + xmlHttp.status);
}
</script>

b.asp
// ============================================================================
<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<?xml version="1.0"  encoding="euc-kr"?>
<response>
  <item>
    <value1>111</value1>
    <value2>특수111+_)(*</value2>
  </item>
  <item>
    <value1>444</value1>
    <value2>특수,./":;'}{][</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>특수$#@!</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>특수^%&amp;</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>한글222&lt;</value2>
  </item>
  <item>
    <value1>333</value1>
    <value2>특수333~`-=\|?</value2>
  </item>
</response>
*/

// ============================================================================
var xmlActionURL = "";
var xmlReturnFunction = "processAjax()";
var xmlErrorFunction = "";
var xmlProcessId = 0;
var xmlHttp;
var xmlDoc;

function initializeURL(url) {
        xmlActionURL = url;
}

function initializeReturnFunction(func) {
        xmlReturnFunction = func;
}

function initializeErrorFunction(errfunc) {
        xmlErrorFunction = errfunc;
}

function createXMLHttpRequest() {
        if (window.ActiveXObject) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        } else if (window.XMLHttpRequest) {
                xmlHttp = new XMLHttpRequest();
        }
}

function doStartRequest(processid) {
        if (xmlProcessId != processid) {
                return;
        }
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", xmlActionURL, true);
        xmlHttp.send(null);
        // alert("" + processid + " 번째 요청처리");
}

function startRequest() {
        // 더블클릭 or <select> 에서의 다량의 요청에 대한 처리
        // 100ms 의 타이머를 두고 해당시간동안 요청이 없을때 실제 요청 시작
        if (xmlProcessId >= 10000) {
                xmlProcessId = 1;
        } else {
                xmlProcessId = xmlProcessId + 1;
        }
        setTimeout("doStartRequest(" + xmlProcessId + ")", 100);
}

function callback() {
        if(xmlHttp.readyState == 4) {
                if(xmlHttp.status == 200) {
                        // 정상적인 데이타 반환
                        // 전체(TXT) : xmlHttp.responseText
                        if (window.ActiveXObject) {
                                // XML 로 변환한다.
                                // 텍스트 앞부분에서 "<" 이전 문자들을 제거한다.(공백문자 제거용,  이렇게 안하면 변환이 안된다 --)
                                xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
                                var rawXML = xmlHttp.responseText;
                                var filteredML;

                                var index = 0;
                                for (var i = 0; i < rawXML.length; i++) {
                                        if (rawXML.charAt(i) == "<") {
                                                index = i;
                                                break;
                                        }
                                }

                                filteredML = rawXML.substring(index);
                                xmlDoc.loadXML(filteredML);
                        } else if (window.XMLHttpRequest) {
                                xmlDoc = xmlHttp.responseXML;
                        }

                        eval(xmlReturnFunction);
                } else if (xmlHttp.status == 204){
                        // 데이터가 존재하지 않을 경우
                        if (xmlErrorFunction == "") {
                                alert("데이타가 존재하지 않습니다.(CODE : " + xmlHttp.status + ")");
                        } else {
                                eval(xmlErrorFunction);
                        }
                } else if (xmlHttp.status == 500){
                        // 에러발생시
                        if (xmlErrorFunction == "") {
                                alert("데이타 수신중 에러가 발생하였습니다.(CODE : " + xmlHttp.status + ")");
                        } else {
                                eval(xmlErrorFunction);
                        }
                }
        }
}