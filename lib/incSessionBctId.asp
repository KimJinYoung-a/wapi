<%

dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

dim C_ADMIN_AUTH
dim C_OFF_AUTH

C_ADMIN_AUTH = (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "kobula") or (session("ssBctId") = "tozzinet") or (session("ssBctId") = "iredfish") or (session("ssBctId") = "kjy8517") or (session("ssBctId") = "okkang77") or (session("ssBctId") = "motions")
C_OFF_AUTH = (session("ssBctId") = "gundolly") or (session("ssBctId") = "hrkang97")


'' 공급업체
dim C_IS_Maker_Upche

'' 가맹점
dim C_IS_OWN_SHOP

'' 직영점
dim C_IS_FRN_SHOP

'' 직영 또는 가맹점
dim C_IS_SHOP

'' 매장 아이디
dim C_STREETSHOPID

''직원
dim C_ADMIN_USER

C_IS_Maker_Upche = (session("ssBctDiv") = "9999")
C_IS_OWN_SHOP = (session("ssBctDiv") = "501") or (session("ssBctDiv") = "502") or (session("ssBctDiv") = "101") or (session("ssBctDiv") = "111") or (session("ssBctDiv") = "112")
''session("ssAdminLsn")=6 매장직원권한 2011-01-14 eastone추가
C_IS_OWN_SHOP = C_IS_OWN_SHOP or (session("ssAdminLsn")="6")

C_IS_FRN_SHOP = (session("ssBctDiv") = "503")
C_IS_SHOP = (C_IS_OWN_SHOP or C_IS_FRN_SHOP)
C_ADMIN_USER     = (session("ssBctDiv") < 10)

if C_IS_FRN_SHOP then
	C_STREETSHOPID = session("ssBctId")
elseif C_IS_OWN_SHOP then
	if (session("ssBctDiv") = "501") or (session("ssBctDiv") = "502") then
		C_STREETSHOPID = session("ssBctid")
	''elseif (session("ssBctDiv")="201" or session("ssAdminPsn")="6") then        ''취화선
	''    C_STREETSHOPID = "cafe002"
	''elseif (session("ssBctDiv")="301" or session("ssAdminPsn")="16") then       ''아카데미
	''    C_STREETSHOPID = "cafe003"
	else
		C_STREETSHOPID = session("ssBctBigo")
	end if
end if

If (session("ssBctId") = "") then
    %><html>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if

'-----------------------------------------------------------------------
' 이벤트 전역변수 선언 (2007.02.07; 정윤정)
'-----------------------------------------------------------------------
Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, fixImgUrl
Dim wwwFingers, imgFingers
''검색엔진 관련
Dim DocSvrAddr, DocSvrPort, DocAuthCode

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'테스트
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
	fixImgUrl			= "http://fiximage.10x10.co.kr"

 	manageUrl 	    = "http://testscm.10x10.co.kr"
 	wwwUrl		    = "http://2010www.10x10.co.kr"            ''차후 정리요망
 	othermall       = "http://othermall.10x10.co.kr"
	mailzine        = "http://testmailzine.10x10.co.kr"
	www2009url      = "http://2009www.10x10.co.kr"
	mobileUrl	    = "http://61.252.133.2"

	wwwFingers		= "http://test.thefingers.co.kr"
	imgFingers		= "http://testimage.thefingers.co.kr"

	''** Upload 구분.;;
	uploadUrl	    = "http://testimgstatic.10x10.co.kr"   ''차후 정리요망
	uploadImgUrl    = "http://testupload.10x10.co.kr"
	ItemUploadUrl	= "http://testupload.10x10.co.kr"
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'웹이미지
		fixImgUrl		= "/fiximage"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		manageUrl 	    = "http://scm.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"

		''** Upload 구분.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "https://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
	else
 		staticImgUrl    = "http://imgstatic.10x10.co.kr"
 		webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지
		fixImgUrl		= "http://fiximage.10x10.co.kr"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		manageUrl 	    = "http://scm.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"

		''** Upload 구분.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
 		ItemUploadUrl	= "http://upload.10x10.co.kr"

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
	end if
END IF

%>
