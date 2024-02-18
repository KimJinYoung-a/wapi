<%

function CheckNChangeCyberAcct(iorderserial)
    dim sqlStr
    dim ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, OLDsubtotalPrice, OLDCancelyn
    ipkumdiv = 0
    OLDsubtotalPrice = 0
    OLDCancelyn      = ""
    
    CheckNChangeCyberAcct = false
    
    sqlStr = " select orderserial, ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice"
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"
    
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        ipkumdiv    = rsget("ipkumdiv")
		accountdiv  = rsget("accountdiv")
		accountNo   = rsget("accountNo")
		cancelyn    = rsget("cancelyn")
		subtotalPrice = rsget("subtotalPrice")
    end if
	rsget.close
	
	if (ipkumdiv<>2) then Exit function
	if (accountdiv<>"7") then Exit function
	
	if (accountNo="���� 470301-01-014754") _
        or (accountNo="���� 100-016-523130") _
        or (accountNo="�츮 092-275495-13-001") _
        or (accountNo="�ϳ� 146-910009-28804") _
        or (accountNo="��� 277-028182-01-046") _
        or (accountNo="���� 029-01-246118") then
            Exit function
    end if
    
    dim CLOSEDATE
    if (cancelyn<>"N") then
        CLOSEDATE = Replace(Left(CStr(now()),10),"-","") & "000000"
    else
        CLOSEDATE = Replace(Left(CStr(DateAdd("d",10,now())),10),"-","") & "235959"
    end if
    
    sqlStr = " select top 1 subtotalPrice, convert(varchar(19),CLOSEDATE,20) as CLOSEDATE "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_CyberAccountLog"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"
    sqlStr = sqlStr & " order by differencekey desc"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        OLDsubtotalPrice = rsget("subtotalPrice")
        OLDCancelyn      = rsget("CLOSEDATE")
        
        if (RIGHT(OLDCancelyn,8)="00:00:00") then
            OLDCancelyn="Y"
        else
            OLDCancelyn="N"
        end if
    end if 
    rsget.close  
    
    if (OLDsubtotalPrice<>subtotalPrice) or (OLDCancelyn<>Cancelyn) then
        CheckNChangeCyberAcct = ChangeCyberAcct(iorderserial, subtotalPrice, CLOSEDATE)
    end if
end function

function ChangeCyberAcct(LGD_OID, LGD_AMOUNT, LGD_CLOSEDATE)
    '/*
    ' * [������� �߱�/�����û ������]
    ' *
    ' * ������� �߱� ����(CHANGE)�� �ݾװ� �����ϸ� ���� �Ҽ� �ֽ��ϴ�. 
    ' */
    dim CST_PLATFORM : CST_PLATFORM         = ""         ' LG�ڷ��� �������� ����(test:�׽�Ʈ, service:����)
    IF application("Svr_Info")="Dev" THEN CST_PLATFORM = "test"
''CST_PLATFORM = ""
    
    dim CST_MID : CST_MID = "tenbyten01"                 ' LG�ڷ������� ���� �߱޹����� �������̵� �Է��ϼ���.
                 
    dim LGD_MID                                                  ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    
    dim LGD_METHOD : LGD_METHOD          = "CHANGE"                              ' ASSIGN:�Ҵ�, CHANGE:����
    
    'LGD_PRODUCTINFO   	 = trim(request("LGD_PRODUCTINFO"))  	 ' ��ǰ����
    'LGD_BUYER          	 = trim(request("LGD_BUYER"))         	 ' �����ڸ�
	'LGD_ACCOUNTOWNER     = trim(request("LGD_ACCOUNTOWNER"))  	 ' �Ա��ڸ�
	'LGD_ACCOUNTPID       = trim(request("LGD_ACCOUNTPID"))       ' �Ա����ֹι�ȣ(�ɼ�)
	'LGD_BUYERPHONE       = trim(request("LGD_BUYERPHONE"))       ' �������޴�����ȣ
	'LGD_BUYEREMAIL       = trim(request("LGD_BUYEREMAIL"))       ' �������̸���(�ɼ�)
	'LGD_BANKCODE         = trim(request("LGD_BANKCODE"))         ' �Աݰ��������ڵ�
	'LGD_CASHRECEIPTUSE   = trim(request("LGD_CASHRECEIPTUSE"))   ' ���ݿ����� ���౸��('1':�ҵ����, '2':��������)
	'LGD_CASHCARDNUM      = trim(request("LGD_CASHCARDNUM"))      ' ���ݿ����� ī���ȣ
	'LGD_TAXFREEAMOUNT    = trim(request("LGD_TAXFREEAMOUNT"))    ' �鼼�ݾ�
	'LGD_CASNOTEURL       = "http://61.252.133.2:8888/admin/apps/DC_CA_noteurl.asp" ''"http://����URL/cas_noteurl.asp"       ' �Աݰ�� ó���� ���� ������������ �ݵ�� ������ �ּ���
	

    'configPath           = "C:/lgdacom"         				 ' LG�ڷ��޿��� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    dim configPath : configPath				   = "C:/lgdacom" '''"C:/lgdacom/conf/" & CST_MID  ''conf ���� ���� 2013/02/15
    
    dim xpay
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    'xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    'xpay.Set "LGD_BUYER", LGD_BUYER
    'xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    'xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    'xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    'xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    'xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    'xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    'xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
    
    'xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    'xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL
    

    '/*
    ' * 1. ������� �߱�/���� ��û ���ó��
    ' *
    ' * ��� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
    ' */
    Dim itemCount, itemName, resCount, i, j
    Dim sqlStr
    
    ChangeCyberAcct = false
    
    if (xpay.TX()) then
        if LGD_METHOD = "ASSIGN" then      '������� �߱��� ���
        
'        	'1)������� �߱ް�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
'        	Response.Write("������� �߱� ��ûó���� �Ϸ�Ǿ����ϴ�. <br>")
'        	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'        	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
'			
'			Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
'	    	Response.Write("�ŷ���ȣ : " & xpay.Response("LGD_TID", 0) & "<p>")
'        	
'        	'�Ʒ��� ������û ��� �Ķ���͸� ��� ��� �ݴϴ�.
'        	
'        	itemCount = xpay.resNameCount
'        	resCount = xpay.resCount
'
'        	For i = 0 To itemCount - 1
'            	itemName = xpay.ResponseName(i)
'            	Response.Write(itemName & "&nbsp:&nbsp")
'            	For j = 0 To resCount - 1
'                	Response.Write(xpay.Response(itemName, j) & "<br>")
'            	Next
'        	Next
        
        else		'������� ������ ���
        	'1)������� ������ ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        	
        	
        	ChangeCyberAcct = (Trim(xpay.resCode)="0000")

        	if (Trim(xpay.resCode)="0000") then
        	    sqlStr = " IF EXISTS (select orderserial from db_order.dbo.tbl_order_CyberAccountLog where orderserial='" & LGD_OID & "')" & VbCrlf
                sqlStr = sqlStr & " BEGIN" & VbCrlf
                sqlStr = sqlStr & "	Insert Into db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	(orderserial, differencekey, userid, FINANCECODE,ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, subtotalPrice, CLOSEDATE"& VbCrlf
                sqlStr = sqlStr & "	,RefIP)" & VbCrlf
                sqlStr = sqlStr & "	select top 1 orderserial, (differencekey+1) as differencekey" & VbCrlf
                sqlStr = sqlStr & "	,userid, FINANCECODE, ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, " & LGD_AMOUNT & " as subtotalprice" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(LGD_CLOSEDATE,4) + "-" + Mid(LGD_CLOSEDATE,5,2) + "-" + Mid(LGD_CLOSEDATE,7,2) + " " + Mid(LGD_CLOSEDATE,9,2) + ":" + Mid(LGD_CLOSEDATE,11,2) + ":" + Mid(LGD_CLOSEDATE,13,2) & "' as CLOSEDATE" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "' as refip" & VbCrlf
                sqlStr = sqlStr & "	from db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	where orderserial='" & LGD_OID & "'" & VbCrlf
                sqlStr = sqlStr & "	order by differencekey desc" & VbCrlf
                sqlStr = sqlStr & " END"
                
                dbget.Execute sqlStr
            ELSE
            	Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
                Response.Write("�ֹ���ȣ : " & LGD_OID & "<br>")
                Response.Write("�Աݾ� : " & LGD_AMOUNT & "<br>")
            	Response.Write("�Աݸ����� : " & LGD_CLOSEDATE & "<p>")
                
                
            	itemCount = xpay.resNameCount
            	resCount = xpay.resCount
            
            	For i = 0 To itemCount - 1
                	itemName = xpay.ResponseName(i)
                	Response.Write(itemName & "&nbsp:&nbsp")
                	For j = 0 To resCount - 1
                    	Response.Write(xpay.Response(itemName, j) & "<br>")
                	Next
            	Next
        	end if
        end if    
    else
        '2)API ��û ���� ȭ��ó��
        ''Response.Write("������� �߱�/���� ��ûó���� ���еǾ����ϴ�. <br>")
        ''Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        ''Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
    
end function

%>