function checkUSBKey ()
{
	var sn='';
	// ActiveX ��ġ���� �˻�
	try {sn = MaGerAuth.GetSN();} catch(e) {alert('���α׷��� ��ġ�Ͽ� �ּ���');}
	// USB Token Ȯ�� ������ �α׾ƿ�
	if(sn=='') {
		top.location = '/login/usbNotFound.asp';
	}

	setTimeout("checkUSBKey()",60000);   //60�ʸ��� �����
}		
