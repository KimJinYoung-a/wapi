<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<%
dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")

''��������
DIM CCOMPID : CCOMPID = session("ssBctId")
''session("ssBctId") = ""

Function is_C_ADMIN_AUTH(loginId)
	Select Case loginId
		'		������		������	   �̼���	 �ϼҶ�	 �̻�	  	�ѿ��	   �����		 ������	   ������		����ȭ	   	������		�̰���			������		  �弮��		����	 ������  		������		������ 		������			������		�豤��			������		ȫ�̼�		�����		������		����		�Ѽ���		������		�赵��		������		��ȿ��
		Case "icommang", "kjy8517", "bseo", "hasora", "skyer9", "tozzinet", "hrkang97", "iredfish", "kobula", "motions",  "thensi7", "llkkjj0906", "angela919", "seokmi1221", "boyishP", "jjy158", "heendoongi", "oesesang52", "hk9566371", "yhj0613", "tenbytendevel", "b00413", "rabbit1693", "sj100", "nys1006", "yooni1105", "celld81", "shj7824", "solleegod", "chh830", "ean0201"
			is_C_ADMIN_AUTH = True
		Case "happy799", "gamt4268", "boom15", "rkgus309"
			is_C_ADMIN_AUTH = True
		Case Else
			is_C_ADMIN_AUTH = False
	End Select
End Function

if (session("ssBctId") = "") then
	if (Request.Cookies("wapi")("UserID") <> "") then

		dim SSO_REGEX, SSO_MATCHES, SSO_LOGIN_SUCCESS, SSO_LOGIN_ID, SSO_LOGIN_HOST, SSO_MATCHES2, SSO_PARTSN
		set SSO_REGEX = new RegExp
		SSO_REGEX.Pattern = "([\w\d]+),([\d]+.[\d]+.[\d]+.[\d]+),([\d]+[-][\d]+[-][\d]+)"
		SSO_REGEX.IgnoreCase = True
		SSO_REGEX.Global = False

		Set SSO_MATCHES = SSO_REGEX.Execute(TBTDecrypt(Request.Cookies("wapi")("UserID")))
		if (SSO_MATCHES.Count > 0) then
			'������ ������ �α����� �����Ƿ� �������� ��ȿ�ϰ� �����Ѵ�.
			if (((Left(now(), 10) = SSO_MATCHES(0).SubMatches(2)) or (Left(DateAdd("d", -1, now()), 10) = SSO_MATCHES(0).SubMatches(2)))) then
				SSO_LOGIN_ID = SSO_MATCHES(0).SubMatches(0)
				SSO_LOGIN_SUCCESS = "Y"
			else
				SSO_LOGIN_SUCCESS = "N"
			end if
		end if

		if (SSO_LOGIN_SUCCESS = "Y") then

			SSO_LOGIN_HOST = Request.ServerVariables("SERVER_NAME")

			if ((SSO_LOGIN_HOST = "testwapi.10x10.co.kr") or (SSO_LOGIN_HOST = "wapi.10x10.co.kr")) then
			    session("ssBctId") = SSO_LOGIN_ID				'ȸ����̵�
			end if
		end if
	end if
end If

DIM CAddDetailSpliter : CAddDetailSpliter= CHR(3)&CHR(4)

dim C_ADMIN_AUTH
C_ADMIN_AUTH = is_C_ADMIN_AUTH(session("ssBctId"))

dim iiisAdmin
iiisAdmin = (session("ssBctId") = "10x10")

if Not iiisAdmin then
	iiisAdmin = (session("ssBctId")<>"")
end if

''2009-10-27 ������ �߰�.
Dim IsAutoScript : IsAutoScript=false

IF (Not iiisAdmin) then
    if (request.Form("redSsnKey")="system") and ((Request.ServerVariables("REMOTE_ADDR")="61.252.133.2") or Request.ServerVariables("REMOTE_ADDR")="110.93.128.99" or (Request.ServerVariables("REMOTE_ADDR")="61.252.133.10") or (Request.ServerVariables("REMOTE_ADDR")="61.252.133.9") or (Request.ServerVariables("REMOTE_ADDR")="110.93.128.94") or (Request.ServerVariables("REMOTE_ADDR")="110.93.128.114") or (Request.ServerVariables("REMOTE_ADDR")="110.93.128.113") or (Request.ServerVariables("REMOTE_ADDR")="61.252.133.70") or (Request.ServerVariables("REMOTE_ADDR")="61.252.133.67") or (Request.ServerVariables("REMOTE_ADDR")="110.93.128.111") or (Request.ServerVariables("REMOTE_ADDR")="121.78.103.60")  ) then
        session("ssBctId")="system"
        session("ssBctDiv")=9
        iiisAdmin = true
        IsAutoScript = true
    end if

	'' eastone
	if ((request.Form("redSsnKey")="system") and ((Request.ServerVariables("REMOTE_ADDR")="222.109.123.95") or (Request.ServerVariables("REMOTE_ADDR")="211.206.236.117"))) then
		session("ssBctId")="system"
        session("ssBctDiv")=9
        iiisAdmin = true
        IsAutoScript = true
    end if

    if ((Request.ServerVariables("REMOTE_ADDR")="165.243.204.101") or (Request.ServerVariables("REMOTE_ADDR")="211.44.122.208") or (Request.ServerVariables("REMOTE_ADDR")="211.44.122.209") or (Request.ServerVariables("REMOTE_ADDR")="192.168.1.67")) then
    	'// GS�� �ֹ��Է� ȣ��
        session("ssBctId")="gseshopapi"
        session("ssBctDiv")=9
        iiisAdmin = true
        IsAutoScript = true
    end if

    '// jenkins
    if Request.ServerVariables("REQUEST_METHOD")= "POST" then
	    if ((request.Form("redSsnKey")="system") and ((Request.ServerVariables("REMOTE_ADDR")="114.31.63.82") or (Request.ServerVariables("REMOTE_ADDR")="172.16.0.225") or (Request.ServerVariables("REMOTE_ADDR")="121.78.103.60"))) then
		    session("ssBctId")="system"
            session("ssBctDiv")=9
            iiisAdmin = true
            IsAutoScript = true
        end if
    else
	    if ((request("redSsnKey")="system") and ((Request.ServerVariables("REMOTE_ADDR")="114.31.63.82") or (Request.ServerVariables("REMOTE_ADDR")="172.16.0.225") or (Request.ServerVariables("REMOTE_ADDR")="121.78.103.60"))) then
		    session("ssBctId")="system"
            session("ssBctDiv")=9
            iiisAdmin = true
            IsAutoScript = true
        end if
    end if
end if

IF (application("Svr_Info")="Dev") and ((request.ServerVariables("REMOTE_ADDR")="::1" or request.ServerVariables("REMOTE_ADDR")="127.0.0.1")) THEN
	'' local �ΰ�� skip
else
	If (Not iiisAdmin) then
	%>
		<script>
		alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.<%=iiisAdmin%>");
		top.location = "/Index.asp";
		</script>
		<%
		response.End
	End if
End if
'-----------------------------------------------------------------------
' �̺�Ʈ �������� ���� (2007.02.07; ������)
'-----------------------------------------------------------------------
 Dim staticImgUrl,uploadUrl,manageUrl,wwwUrl, uploadImgUrl,othermall,mailzine,www2009url, ItemUploadUrl, staticUploadUrl, webImgUrl, mobileUrl, partnerUrl, fixImgUrl
 Dim vwwwUrl, vmobileUrl
 Dim wwwFingers, imgFingers, wwwithinksoweb, wwwithinkso, UploadDefaultPath
  ''�˻����� ����
 Dim DocSvrAddr, DocSvrPort, DocAuthCode

IF application("Svr_Info")="Dev" THEN
 	staticImgUrl = "http://testimgstatic.10x10.co.kr"	'�׽�Ʈ
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���
	fixImgUrl		= "http://fiximage.10x10.co.kr"

 	manageUrl 	    = "http://testscm.10x10.co.kr"
 	wwwUrl		    = "http://2010www.10x10.co.kr"            ''���� �������
 	vwwwUrl			= "http://2013www.10x10.co.kr"
 	othermall       = "http://othermall.10x10.co.kr"
	mailzine        = "http://testmailzine.10x10.co.kr"
	www2009url      = "http://2009www.10x10.co.kr"
	mobileUrl	    = "http://2013m.10x10.co.kr"
	vmobileUrl	    = "http://2013m.10x10.co.kr"

	wwwFingers		= "http://test.thefingers.co.kr"
	imgFingers		= "http://testimage.thefingers.co.kr"
	wwwithinkso		= "http://devwww.ithinkso.co.kr"
	wwwithinksoweb  = "http://test.ithinksoweb.com"

	''** Upload ����.;;
	uploadUrl	    = "http://testimgstatic.10x10.co.kr"   ''���� �������
	uploadImgUrl    = "http://testupload.10x10.co.kr"
	ItemUploadUrl	= "http://testupload.10x10.co.kr"
	partnerUrl		= "http://testwebimage.10x10.co.kr/partner"		'�ӽû�ǰ�̹���(��Ʈ��)
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		staticImgUrl    = "/imgstatic"
 		webImgUrl		= "/webimage"							'���̹���
		fixImgUrl		= "/fiximage"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		vwwwUrl 		= "http://www.10x10.co.kr"
 		manageUrl 	    = "http://scm.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m1.10x10.co.kr"
		vmobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"
		wwwithinkso		= "http://www.ithinkso.co.kr"
		wwwithinksoweb  = "http://www.ithinksoweb.com"

		''** Upload ����.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "https://upload.10x10.co.kr"
 		partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
	else
 		staticImgUrl    = "http://imgstatic.10x10.co.kr"
 		webImgUrl		= "http://webimage.10x10.co.kr"				'���̹���
		fixImgUrl		= "http://fiximage.10x10.co.kr"

 		wwwUrl 		    = "http://www1.10x10.co.kr"
 		vwwwUrl 		= "http://www.10x10.co.kr"
 		manageUrl 	    = "http://scm.10x10.co.kr"
 		othermall       = "http://gseshop.10x10.co.kr"
		mailzine        = "http://mailzine.10x10.co.kr"
		www2009url      = "http://www.10x10.co.kr"
		mobileUrl	    = "http://m1.10x10.co.kr"
		vmobileUrl	    = "http://m.10x10.co.kr"

		wwwFingers		= "http://www.thefingers.co.kr"
		imgFingers		= "http://image.thefingers.co.kr"
		wwwithinkso		= "http://www.ithinkso.co.kr"
		wwwithinksoweb  = "http://www.ithinksoweb.com"

		''** Upload ����.;;
		uploadUrl	    = "http://oimgstatic.10x10.co.kr"
		uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
 		ItemUploadUrl	= "http://upload.10x10.co.kr"
 		partnerUrl		= "http://partner.10x10.co.kr"				'�ӽû�ǰ�̹���(��Ʈ��)

		staticUploadUrl = "http://oimgstatic.10x10.co.kr"
	end if
END IF
%>
