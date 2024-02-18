<%@ language=vbscript %>
<% option explicit %>
<%
'isTest => 웹에서 테스트 하려면 True로 할 것, App에서 사용하려면 False
Dim isTest : isTest = False
If isTest <> True Then
	Response.ContentType = "application/json"
	Response.AddHeader "Accept", "application/json"
	Response.Charset = "UTF-8"
Else
	Response.Charset = "UTF-8"
End If
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/JSON_noenc.2.0.4.asp"-->
<!-- #include virtual="/syrup/syrupCommFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
If (Now() > #03/31/2017 14:00:00#) Then
'	response.write "Syrup 시스템 작업 중 입니다"
	response.end	
End If

Dim sData, oResult, sqlStr, errCode
'Dim patrn : patrn = "(""member_pw"":"")[\s\S]*("")"
Dim ret, refip

'Json데이터를 바이너리로 받아서 텍스트 형식으로 읽는다.
If isTest <> True Then
	On Error Resume Next
		sData = BinaryToText(Request.BinaryRead(request.TotalBytes), "UTF-8")
		If sData <> "" Then
			sData = URLDecodeUTF8(sData)
		End If

		If (Err) Then
			response.write "JSON 데이터 호출 오류"
			response.end
		End If
	On Error Goto 0
Else
	'sData= "{""cd_encryption_key"":""1"",""data"":{""birthday"":""19820505"",""agree_count"":""3"",""cd_sex"":""M"",""cust_name"":""한용민"",""pay_trace_num"":""40044330"",""member_pw"":""qpqp1010"",""eng_first_name"":"""",""cd_comm_agency"":""0101"",""pay_approve_num"":"""",""jumin_num"":""8205051"",""flag_pay_complete"":""Y"",""pay_amount"":"""",""cd_reg_channel"":""1"",""ipin_num"":"""",""flg_road_addr"":""N"",""flag_foreigner"":""K"",""cd_reg_shop"":""SWMT"",""hp_num"":""01091778708"",""agree_data"":[{""agree_value"":""Y"",""agree_code"":""1""},{""agree_value"":""Y"",""agree_code"":""2""},{""agree_value"":""Y"",""agree_code"":""3""}],""post_num"":""132020"",""member_id"":""test0011"",""addr1"":""서울 도봉구 방학동"",""addr2"":""123-123"",""email"":"""",""ci"":""gGlOzbyn6uS1txJ2GTVl9ukBWdiTxR7dbT07C6sQD00p1XJO0D2pKOYHG2E5iplxjs\/HhvalIrViQ6pv80agRg=="",""pay_product_name"":"""",""eng_last_name"":""""},""seq_trans"":""20150206183831274653"",""cd_fulltext"":""1100"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"

	'아이디는 안 넘어오고 ci값이 텐바이텐CI와 같을 때
	'sData= "{""cd_encryption_key"":""1"",""data"":{""birthday"":""19850101"",""agree_count"":""3"",""cd_sex"":""F"",""cust_name"":""테스터"",""pay_trace_num"":"""",""member_pw"":"""",""eng_first_name"":"""",""cd_comm_agency"":""0101"",""pay_approve_num"":"""",""jumin_num"":""8501012"",""flag_pay_complete"":"""",""pay_amount"":""0000000000"",""cd_reg_channel"":"""",""ipin_num"":""MC0GCCqGSIb3DQIJAyEAKpiXMIPdihcwXn8yQ9pjzIM\/wDGSgVGYH0IvE3RspTs="",""flg_road_addr"":""N"",""flag_foreigner"":""K"",""cd_reg_shop"":"""",""hp_num"":""01012345678"",""agree_data"":[{""agree_value"":""Y"",""agree_code"":""1""},{""agree_value"":""Y"",""agree_code"":""2""},{""agree_value"":""Y"",""agree_code"":""3""}],""post_num"":"""",""member_id"":"""",""addr1"":""서울"",""addr2"":""강남구"",""email"":"""",""ci"":""t+BDMoutg9vTX6Wx3V6MrcqQ03I1e2L8soerevg3arGsZWr00FuSEgqn8jS5QdZFH8TfyspxZstdd55faVa29A=="",""pay_product_name"":"""",""eng_last_name"":""""},""seq_trans"":""00000000000000000002"",""cd_fulltext"":""1100"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"

	'로그인 유도
	'sData= "{""cd_encryption_key"":""1"",""data"":{""birthday"":""19850101"",""agree_count"":""3"",""cd_sex"":""F"",""cust_name"":""테스터"",""pay_trace_num"":"""",""member_pw"":"""",""eng_first_name"":"""",""cd_comm_agency"":""0101"",""pay_approve_num"":"""",""jumin_num"":""8501012"",""flag_pay_complete"":"""",""pay_amount"":""0000000000"",""cd_reg_channel"":"""",""ipin_num"":""MC0GCCqGSIb3DQIJAyEAKpiXMIPdihcwXn8yQ9pjzIM\/wDGSgVGYH0IvE3RspTs="",""flg_road_addr"":""N"",""flag_foreigner"":""K"",""cd_reg_shop"":"""",""hp_num"":""01012345678"",""agree_data"":[{""agree_value"":""Y"",""agree_code"":""1""},{""agree_value"":""Y"",""agree_code"":""2""},{""agree_value"":""Y"",""agree_code"":""3""}],""post_num"":"""",""member_id"":"""",""addr1"":""서울"",""addr2"":""강남구"",""email"":"""",""ci"":""t+BDMoutg9vTX6Wx3V6MrcqQ03I1e2L8soerevg3arGsZWr00FuSEgqn8jS5QdZFH8TfyspxZstdd55faVa29A="",""pay_product_name"":"""",""eng_last_name"":""""},""seq_trans"":""00000000000000000002"",""cd_fulltext"":""1100"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"

	'로그인해서 넘어왔을 때
	'sData= "{""cd_encryption_key"":""1"",""data"":{""birthday"":""19850101"",""agree_count"":""3"",""cd_sex"":""F"",""cust_name"":""테스터"",""pay_trace_num"":"""",""member_pw"":"""",""eng_first_name"":"""",""cd_comm_agency"":""0101"",""pay_approve_num"":"""",""jumin_num"":""8501012"",""flag_pay_complete"":"""",""pay_amount"":""0000000000"",""cd_reg_channel"":"""",""ipin_num"":""MC0GCCqGSIb3DQIJAyEAKpiXMIPdihcwXn8yQ9pjzIM\/wDGSgVGYH0IvE3RspTs="",""flg_road_addr"":""N"",""flag_foreigner"":""K"",""cd_reg_shop"":"""",""hp_num"":""01012345678"",""agree_data"":[{""agree_value"":""Y"",""agree_code"":""1""},{""agree_value"":""Y"",""agree_code"":""2""},{""agree_value"":""Y"",""agree_code"":""3""}],""post_num"":"""",""member_id"":""kjy8517"",""addr1"":""서울"",""addr2"":""강남구"",""email"":"""",""ci"":""t+BDMoutg9vTX6Wx3V6MrcqQ03I1e2L8soerevg3arGsZWr00FuSEgqn8jS5QdZFH8TfyspxZstdd55faVa29A=="",""pay_product_name"":"""",""eng_last_name"":""""},""seq_trans"":""00000000000000000002"",""cd_fulltext"":""1100"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"

	'시럽 신규가입 일때
	'sData= "{""cd_encryption_key"":""1"",""data"":{""birthday"":""19850101"",""agree_count"":""3"",""cd_sex"":""F"",""cust_name"":""테스터"",""pay_trace_num"":"""",""member_pw"":""test0000"",""eng_first_name"":"""",""cd_comm_agency"":""0101"",""pay_approve_num"":"""",""jumin_num"":""8501012"",""flag_pay_complete"":"""",""pay_amount"":""0000000000"",""cd_reg_channel"":"""",""ipin_num"":""MC0GCCqGSIb3DQIJAyEAKpiXMIPdihcwXn8yQ9pjzIM\/wDGSgVGYH0IvE3RspTs="",""flg_road_addr"":""N"",""flag_foreigner"":""K"",""cd_reg_shop"":"""",""hp_num"":""01012345678"",""agree_data"":[{""agree_value"":""Y"",""agree_code"":""1""},{""agree_value"":""Y"",""agree_code"":""2""},{""agree_value"":""Y"",""agree_code"":""3""}],""post_num"":"""",""member_id"":""test0000"",""addr1"":""서울"",""addr2"":""강남구"",""email"":"""",""ci"":""fyspxZstdd55faVa29A=="",""pay_product_name"":"""",""eng_last_name"":""""},""seq_trans"":""00000000000000000002"",""cd_fulltext"":""1100"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"
End If
'################## 공통 변수 ###################
Dim jcd_fulltext, jcd_partner, jseq_trans
'################# Data Arr변수 #################
Dim jbirthday, jagree_count, jcd_sex, jcust_name, jpay_trace_num, jmember_pw, jeng_first_name, jcd_comm_agency, jpay_approve_num, jjumin_num, jflag_pay_complete, jpay_amount, jcd_reg_channel
Dim jipin_num, jflg_road_addr, jflag_foreigner, jcd_reg_shop, jhp_num, jpost_num, jmember_id, jaddr1, jaddr2, jemail, jci, jpay_product_name, jeng_last_name
Dim jsmsok, jemailok, repData
'################################################
On Error Resume Next
SET oResult = JSON.parse(sData)
	jcd_fulltext		= oResult.cd_fulltext							'사용 코드
	jcd_partner			= oResult.cd_partner							'고객사 코드
	jseq_trans			= oResult.seq_trans								'추적번호

	jbirthday			= oResult.data.birthday							'생년월일
	jagree_count		= oResult.data.agree_count						'약관 개수
	jcd_sex				= oResult.data.cd_sex							'성별
	jcust_name			= oResult.data.cust_name						'고객명
	jpay_trace_num		= oResult.data.pay_trace_num					'결제 추적번호
	jmember_pw			= oResult.data.member_pw						'회원 비밀번호
	jeng_first_name		= oResult.data.eng_first_name					'영문 이름
	jcd_comm_agency		= oResult.data.cd_comm_agency					'통신사 코드
	jpay_approve_num	= oResult.data.pay_approve_num					'결제 승인번호
	jjumin_num			= oResult.data.jumin_num						'주민번호7자리
	jflag_pay_complete	= oResult.data.flag_pay_complete				'결제 완료여부
	jpay_amount			= oResult.data.pay_amount						'결제 금액
	jcd_reg_channel		= oResult.data.cd_reg_channel					'가입경로(채널)코드
	jipin_num			= oResult.data.ipin_num							'I-PIN DI
	jflg_road_addr		= oResult.data.flg_road_addr					'도로명주소 사용여부
	jflag_foreigner		= oResult.data.flag_foreigner					'외국인 구분 코드 | A:외국인, K:내국인
	jcd_reg_shop		= oResult.data.cd_reg_shop						'가입매장 코드
	jhp_num				= oResult.data.hp_num							'휴대폰번호
	jpost_num			= oResult.data.post_num							'우편번호
	jmember_id			= oResult.data.member_id						'회원 ID
	jaddr1				= oResult.data.addr1							'주소1
	jaddr2				= oResult.data.addr2							'주소2
'	jemail				= oResult.data.email							'이메일
	jci					= oResult.data.ci								'CI
	jpay_product_name	= oResult.data.pay_product_name					'결제 상품명
	jeng_last_name		= oResult.data.eng_last_name					'영문 성

	jsmsok 				= oResult.data.agree_data.get(2).agree_value	'약관화면에서 SMS수신여부
	If (Err) Then
		'해당되는 json데이터가 없거나 하는 이유로 에러라면 에러메세지 출력
		errCode = "2101"
	Else
		If jci = "" Then
			errCode = "2107"
		ElseIf year(now()) - Left(jbirthday, 4) < 14 Then
			errCode = "2103"
		Else
			errCode = "0000"
		End If
	End If
SET oResult = nothing

'// 진영 추가 refip 및 json 정보치환//////////////////////////////////////////////////////////////////
refip	= request.ServerVariables("REMOTE_ADDR")
repData = ""
repData = repData & "cust_name : " & jcust_name
repData = repData & ", cd_reg_channel : " & jcd_reg_channel
repData = repData & ", ipin_num : " & jipin_num
repData = repData & ", hp_num : " & jhp_num
repData = repData & ", ci : " & jci
repData = repData & ", seq_trans : " & jseq_trans
repData = repData & ", cd_fulltext : " & jcd_fulltext
repData = repData & ", cd_partner : " & jcd_partner
repData = html2db(repData)
'///////////////////////////////////////////////////////////////////////////////////////////////////
On Error Goto 0
'################################ 여기 까지 왔다면 적어도 JSON 데이터가 왔다는 것 #############################
If (jmember_pw = "") Then
	sqlStr = ""
	'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('A', 'cardsend', getdate(), '"&sData&"', '"&refip&"') "
	sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('A', 'cardsend', getdate(), '"&repData&"', '"&refip&"') "
	dbCTget.execute sqlStr
Else
	Dim passPatrn : passPatrn = """member_pw"":"""&jmember_pw&""""
	If Instr(sData, passPatrn) > 0 Then
		ret = replace(sData, passPatrn, """member_pw"":""""")
	End If

'	ret = RepWord(sData, patrn, """member_pw"":""""")
	sqlStr = ""
	'sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('A', 'cardsend', getdate(), '"&ret&"', '"&refip&"') "
	sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('A', 'cardsend', getdate(), '"&repData&"', '"&refip&"') "
	dbCTget.execute sqlStr
End If
'##############################################################################################################

'############################################### 하단 If문 설명 ###############################################
'====== jmember_id가 오지 않았다는 것은 아래로 설명됨 =======
'1.시럽 재설치
'2.시럽은 처음 설치했으나 텐바이텐 회원(CI로 체크)
'====== jmember_id가 왔다는 것은 아래로 설명됨		  =======
'1.로그인 완료
'2.신규회원으로 jmember_pw까지 동반해서 옴
'##############################################################################################################
If (jmember_id = "") Then			'로그인 화면 나오기 전 약관동의했을 때
	'회원ID가 오지 않았을 때는 우선 넘어온 CI를 우리쪽 CI와 비교
	If isMemberToCI(jci, jmember_id) Then
		'넘어온 CI와 우리쪽 CI가 같기 때문에 있으면 회원으로 간주하고 카드발급 및 각 테이블 인서트 안한것 인서트
		Call fnShopUserJsonFlushProc(jseq_trans, jmember_id, jcust_name, jhp_num, jci, errCode)
		'#######################쿠폰 기간동안 신규가입자 쿠폰 증정############################
		If (Now() > #06/08/2015 00:00:00# AND Now() < #07/07/2015 23:59:59#) Then				'---- 1번째
			Call isNewCardCoupon(jmember_id)
		End If
		'#####################################################################################
		Call fnConnInfoUpdateProc(jci, jmember_id)		'CI값 user_n과 tbl_total_shop_user에 저장
	Else
		'넘어온 CI와 우리쪽 CI가 다르기 때문에 로그인 유도
		Call fnLoginJsonFlush(jseq_trans, jcust_name, jhp_num, errCode)
	End If
Else								'로그인성공 또는 회원가입일 때
	If (jmember_pw = "") Then
		'기존 텐바이텐 회원
		Call fnShopUserJsonFlushProc(jseq_trans, jmember_id, jcust_name, jhp_num, jci, errCode)
		'#######################쿠폰 기간동안 신규가입자 쿠폰 증정############################
		If (Now() > #06/08/2015 00:00:00# AND Now() < #07/07/2015 23:59:59#) Then				'---- 2번째
			Call isNewCardCoupon(jmember_id)
		End If
		'#####################################################################################
		Call fnConnInfoUpdateProc(jci, jmember_id)		'CI값 user_n과 tbl_total_shop_user에 저장
	Else
		'시럽에서 신규가입
		Call fnJoin10x01Flush(jseq_trans, jbirthday, jagree_count, jcd_sex, jcust_name, jmember_pw, jjumin_num, jhp_num, jmember_id, jci, jsmsok, errCode)
		'#######################쿠폰 기간동안 신규가입자 쿠폰 증정############################
		If (Now() > #06/08/2015 00:00:00# AND Now() < #07/07/2015 23:59:59#) Then				'---- 3번째
			Call isNewCardCoupon(jmember_id)
		End If
		'#####################################################################################
	End If
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->