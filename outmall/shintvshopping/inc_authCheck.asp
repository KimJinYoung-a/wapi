<%
Dim shintvshoppingAPIURL, linkCode, entpCode, entpId, entpPass, shipCostCode, mdCode, entpManSeq, returnManSeq, shipManSeq, makecoCode, originCode, brandCode, mdManId

IF application("Svr_Info") = "Dev" THEN
	shintvshoppingAPIURL = "http://open-api-dev.shinsegaetvshopping.com"
	linkCode		= "TENBY"			'�����ڵ�
	entpCode		= "410000"			'��ü�ڵ�
	entpId			= "E410000"			'��ü�����ID
	entpPass		= "E410000"			'��üPASSWORD
	entpManSeq		= "002"				'��ü�����
	returnManSeq	= "004"				'ȸ�������
	shipManSeq		= "003"				'�������
	shipCostCode	= "B001"			'��ۺ���å�ڵ� | 5�����̻� 3õ��
	mdCode 			= "061"				'MD
	mdManId			= "KYMMD"			'���MD ID
	makecoCode		= "AES2"			'������ü | ���߼��������� "�󼼼�������"
	originCode		= "9999"			'������ | ���߼��������� "�󼼼�������"
	brandCode		= "022652"			'�귣�� | test
Else
	Dim shintvshoppingStrSql
	shintvshoppingStrSql = ""
	shintvshoppingStrSql = shintvshoppingStrSql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
	shintvshoppingStrSql = shintvshoppingStrSql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
	shintvshoppingStrSql = shintvshoppingStrSql & " where mallid='shintvshopping' " & VbCRLF
	shintvshoppingStrSql = shintvshoppingStrSql & " and inikey='pass'"
	rsget.CursorLocation = adUseClient
	rsget.Open shintvshoppingStrSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		entpPass	= rsget("iniVal")
	end if
	rsget.close

	shintvshoppingAPIURL = "http://open-api.shinsegaetvshopping.com"
	linkCode		= "TENBY"			'�����ڵ�
	entpCode		= "419803"			'��ü�ڵ�
	entpId			= "E419803"			'��ü�����ID
'	entpPass		= "ten101010*"		'��üPASSWORD
	entpManSeq		= "001"				'��ü����� | ������
	returnManSeq	= "005"				'ȸ������� | ������
	shipManSeq		= "004"				'������� | ������
	shipCostCode	= "B01"				'��ۺ���å�ڵ� | 5�����̻� ������
	mdCode 			= "061"				'MD | 061 : �¶���
	mdManId			= "011074"			'���MD ID
	makecoCode		= "AES2"			'������ü | �󼼼�������
	originCode		= "9999"			'������ | �󼼼�������
	brandCode		= "031506"			'�귣��
End if
%>