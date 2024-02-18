<%@ language=vbscript %>
<% option explicit %>
<%
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

Dim sData, oResult, sqlStr, errCode, refip
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
	'sData= "{""cd_encryption_key"":""1"",""data"":{""ci"":""111"", ""cust_num"":""222"", ""cust_name"":""홍길도"",""card_grade"":""5"",""card_num"":""1010990070253380"",""hp_num"":""01011111111""},""seq_trans"":""00000000000000000003"",""cd_fulltext"":""1210"",""cd_partner"":""642"",""cd_encryption"":""00"",""version"":""0021""}"
	sData= "{""cd_encryption_key"":""1"",""data"":{""cust_name"":"""",""card_grade"":""006"",""cust_num"":"""",""ci"":""t+BDMoutg9vTX6Wx3V6MrcqQ03I1e2L8soerevg3arGsZWr00FuSEgqn8jS5QdZFH8TfyspxZstdd55faVa29A=="",""card_num"":""1010990020720724"",""hp_num"":""""},""seq_trans"":""00000000000000000008"",""cd_fulltext"":""1200"",""cd_partner"":""971"",""cd_encryption"":""00"",""version"":""0031""}"
End If
'################## 공통 변수 ###################
Dim jcd_fulltext, jcd_partner, jseq_trans
'################# Data Arr변수 #################
Dim jcard_num, jhp_num, jcust_name, jci, jcust_num, jcard_grade
'################################################
On Error Resume Next
SET oResult = JSON.parse(sData)
	jcd_fulltext		= oResult.cd_fulltext						'사용 코드
	jcd_partner			= oResult.cd_partner						'고객사 코드
	jseq_trans			= oResult.seq_trans							'추적번호

	jcard_num			= oResult.data.card_num						'카드번호
	jhp_num				= oResult.data.hp_num						'휴대폰번호
	jcust_name			= oResult.data.cust_name					'고객명
	jci					= oResult.data.ci							'CI
	jcust_num			= oResult.data.cust_num						'고객관리번호
	jcard_grade			= oResult.data.card_grade					'카드 등급 코드

	If (Err) Then
		errCode = "3101"
	Else
		errCode = "0000"
	End If
SET oResult = nothing
refip = request.ServerVariables("REMOTE_ADDR")
On Error Goto 0

'################################ 여기 까지 왔다면 적어도 JSON 데이터가 왔다는 것 #############################
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('V', 'pointDetailview', getdate(), '"&sData&"', '"&refip&"') "
dbCTget.execute sqlStr
'##############################################################################################################
Call fnpointDetailViewFlush(jseq_trans, jcust_name, jhp_num, jcard_num, jcust_num, jcard_grade, errCode)

sqlStr = ""
sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_syrupCheckLog (chk_successYN, ref, regdate, jdata, refip) VALUES ('V', 'pointDetailview_Fin', getdate(), '"&sData&"', '"&refip&"') "
dbCTget.execute sqlStr

If jcust_num <> "" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_shop.dbo.tbl_total_shop_user SET "
	sqlStr = sqlStr & " lastQueDate = getdate() "
	sqlStr = sqlStr & " WHERE UserSeq = '"&jcust_num&"' "
	dbget.Execute sqlStr, 1
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->