<%
'인증,결제 및 웹 경로
Const KMPAY_WEB_SERVER_URL = "https://kmpay.lgcns.com:8443"
Const KMPAY_CERT_SERVER_URL = "https://kmpay.lgcns.com:8443"
Const KMPAY_CERT_SERVER_PAGE = "/merchant/requestDealApprove.dev"
Const CNSPAY_DEAL_REQUEST_URL = "https://pg.cnspay.co.kr:443"

'TXN_ID 호출전용 키값
DIM KMPAY_MERCHANT_ID : KMPAY_MERCHANT_ID= "KCTEN0001m"
DIM KMPAY_MERCHANT_ENCKEY : KMPAY_MERCHANT_ENCKEY= "689dcff525811766"
DIM KMPAY_MERCHANT_HASHKEY : KMPAY_MERCHANT_HASHKEY= "7195bb5277817133"

'가맹점서명키 (꼭 해당 가맹점키로 바꿔주세요)
DIM KMPAY_MERCHANT_KEY : KMPAY_MERCHANT_KEY= "+f41Tt/PW/EtEuUEf0kYBpK4D41Jtpt3l9LP+Fke1wYS8nv1KL9+N0b8qnHecxMmyIMCasagGApBNAGSCuCuHw=="

'취소 비밀번호
Dim KMPAY_CANCEL_PWD : KMPAY_CANCEL_PWD= "KCTEN0001" ''다른걸로 하믄 취소시 안됨?..

if (application("Svr_Info")="Dev") then
    KMPAY_MERCHANT_ID = "cnstest25m"
    KMPAY_MERCHANT_ENCKEY = "10a3189211e1dfc6"
    KMPAY_MERCHANT_HASHKEY = "10a3189211e1dfc6"
    
    '가맹점서명키 (꼭 해당 가맹점키로 바꿔주세요)
    KMPAY_MERCHANT_KEY = "33F49GnCMS1mFYlGXisbUDzVf2ATWCl9k3R++d5hDd3Frmuos/XLx8XhXpe+LDYAbpGKZYSwtlyyLOtS/8aD7A=="
    
    '취소 비밀번호
    KMPAY_CANCEL_PWD = "123456" ''다른걸로 하믄 취소시 안됨?..
end if

'로그
Const KMPAY_LOG_DIR = "C:/KMPay/Log/"
Const KMPAY_LOG_LEVEL = 2   '-1:로그 사용 안함, 0:Error, 1:Info, 2:Debug  //
%>