<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1200  ''�ʴ���
'��ǰEP�� 78�� DB�� �ٶ󺸰�, �Ǹ�EP�� 77��DB�� �ٶ󺻴�
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' ���̹� ���ļ��� ���� Make / �Ϻ�
Const MaxPage   = 300   ''maxpage ���� 40->50���� 2013-12-13����, 50->60���� 2014-09-23 ������ ����, 60->70���� 2014-10-08 ���� ,70->100 ���� 2016-06-29
Const PageSize = 5000  ''3000->5000

Dim appPath : appPath = server.mappath("/Files/naverEP/") + "\"
Dim FileName: FileName = "naverNewVerDailyEP_temp.txt"
Dim newFileName: newFileName = "naverNewVerDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "naverNewVerChangedEP_temp.txt"
	newFileName = "naverNewVerChangedEP.txt"
End If

Function WriteMakeNaverFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr, isMake, basicImage, basic600Image, displayImageUrl
    Dim itemid,deliverytype, deliv, dispCash
    Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname, evtText, isCouponDown, nvcpnVal, iNvCouponPro, iNvCouponValue, deliveryFixday, importFlagYN, adultType
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
'���ϴ� ����ī�װ�
		displayImageUrl = ""
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''��ۺ� /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 �߰�
		    ArrCateNM		= ""
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= ""
    		CntCD			= Split(ArrCateCD,",")
    		jaehu3depNM		= ""
		else
    		ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
    		CntCD			= Split(ArrCateCD,",")
    		jaehu3depNM		= Split(arrList(20,intLoop),"||")(2)

    		'2������ 2�������� ������ ����..2017-10-17 ������
    		If Ubound(CntNM) = 1 then
				jaehu3depNM = Split(ArrCateNM, ",")(1)
	    	End If
        end if

		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		basicImage		= arrList(4,intLoop)
		basic600Image	= arrList(34,intLoop)

		If basic600Image <> "" Then
			displayImageUrl = "http://webimage.10x10.co.kr/image/basic600/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop)
		Else
			displayImageUrl = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop)
		End If

		If itemid = "1831400" then	''2017-12-01 ���µ� ����� ��û..��ǰ�� ���� ����
			itemname = "1.1M �����̾� ������ ũ��������Ʈ�� ����Ǯ��Ʈ (����)_(540048)_Ʈ��"
		End If

		If (deliverytype = "7") Then deliv=-1
		If arrList(27,intLoop) = "06" OR arrList(27,intLoop) = "16" Then
			isMake = "Y"
		Else
			isMake = "N"
		End If

		If arrList(28,intLoop) > 0 Then					'���̹� �������� ������...����Ȯ���Ͽ� ���������� nvcpnVal�� ���� �����ؾ���
			dispCash	= CLNG(arrList(28,intLoop))

			'' �ּ�ó�� 2019/05/20
			' iNvCouponPro = CLNG(arrList(29,intLoop))  ''2018/03/09 �߰�
			' iNvCouponValue = CLNG(arrList(30,intLoop))  ''2018/03/23 �߰�

			' If iNvCouponValue > 0 Then
			' 	evtText		= "�ڳ��̹����� �߰����Ρ�"
			' 	isCouponDown= "Y"
			' 	nvcpnVal	= Replace(arrList(22,intLoop),"&nbsp;","")
			' Else
			' 	if (iNvCouponPro>0) and (iNvCouponPro<100) then  ''2018/03/09 ����
	    	' 		evtText		= "�ڳ��̹����� "&iNvCouponPro&"% �߰����Ρ�"
	    	' 		isCouponDown= "Y"
	    	' 		nvcpnVal	= "^"&iNvCouponPro   ''���������� ��� 1~99 ���� ���� (%)
			'     end if
			' End If
		Else
			dispCash	= CLNG(arrList(3,intLoop))

			'' �ּ�ó�� 2019/05/20
			' If (FALSE) AND (Now() > #10/13/2017 00:00:00# AND Now() < #10/25/2017 20:59:59#) Then  ''��¥ ���� ��/��/�⵵
			' 	evtText		= "�ٹ����� 16�ֳ� ������! �ִ� 30% ��������"
			' ELSEIF (Now() > #10/01/2018 00:00:00# AND Now() < #10/01/2018 21:59:59#) Then
			' 	evtText		= "10/1�� ���� �� �Ϸ縸! �ִ� 3���� ��������"
			' Else
			' 	evtText		= "�� ���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"
			' End If

			' isCouponDown= ""
			' nvcpnVal	= ""
		End If

		'' �̺�Ʈ ���� ���� 2019/05/20
		'' �̺�Ʈ ���� DBȭ 2019-09-25 ������ �߰�
		evtText		= arrList(33,intLoop)
		isCouponDown= ""
		nvcpnVal	= ""

		'�켱 ���� Depth3ItemName > Depth3MakerName > ����ī�װ���
		If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
			IF (isIsChangedEP) then			'���EP
				If arrList(21,intLoop) = "U" Then	'��������(U)
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			Else
				If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			End If
		End If

		deliveryFixday = arrList(31,intLoop)
		If deliveryFixday = "G" Then
			importFlagYN = "Y"
		Else
			importFlagYN = ""
		End If

		''2019/04/25
		adultType = arrList(32,intLoop)
		if (adultType="1" or adultType="2") then
			adultType="Y"
		else
			adultType=""
		end if


        '' 2018/03/09 "_" => CHKIIF(jaehu3depNM="",""," ")
		bufstr = itemid & vbTab & Replace(itemname, vbTab, "") & CHKIIF(jaehu3depNM="",""," ") & jaehu3depNM & vbTab & dispCash & vbTab & dispCash & vbTab  		'��ǰ�ڵ� | ��ǰ�� | pc�ǸŰ��� | ����� �ǸŰ���
'2019-04-11 �ϴ� �ּ�ó��
'		bufstr = bufstr & CLNG(arrList(26,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=nvshop_sp&utm_source=organic&utm_medium=shopping_w&utm_campaign=nvshop_w&term=nvshop" & vbTab	'���� | ��ǰURL
'		bufstr = bufstr & "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=nvshop_sp&utm_source=organic&utm_medium=shopping_m&utm_campaign=nvshop_m&term=nvshop" & vbTab									'��ǰ�����URL
'2019-04-11 ���ú��ش� ��û �ϴ����� utmParam ����
		bufstr = bufstr & CLNG(arrList(26,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_w&term=nvshop_w&rdsite=nvshop_sp" & vbTab	'���� | ��ǰURL
		'// ������� �귣ġ�� ����
		bufstr = bufstr & "http://m.10x10.co.kr/common/tenlanding.asp?urltype=item&itemid="&itemid&"&utm_source=naver&utm_medium=organic&utm_campaign=shopping_m&term=nvshop_m&rdsite=nvshop_sp" & vbTab									'��ǰ�����URL

'����ǰ�� �Ʒ� ���ǹ����� ���� �ʰ� GetImageSubFolderByItemid �����Ͽ� ���� / 2020-01-21 ������ ����
'		if (arrList(27,intLoop)="21") then
'		bufstr = bufstr & "http://webimage.10x10.co.kr/image/basic/" & arrList(4,intLoop) & vbTab & "" & vbTab	'�̹���URL | �߰� �̹���URL
'		else
		bufstr = bufstr & displayImageUrl & vbTab & "" & vbTab	'�̹���URL | �߰� �̹���URL
'		end if

		For lp = 1 to Ubound(CntNM) + 1
			If lp>4 Then Exit For
			bufstr = bufstr & Replace(CntNM(lp-1),"&nbsp;","") & vbTab																						'���޻� ī�װ���(��/��/��/��)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr & "" & vbTab
			Next
		End If

		if (itemid="2142647") then  ''���θ����׽�Ʈ
		bufstr = bufstr & "" & vbTab & "15883309361" & vbTab & "�Ż�ǰ" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'���̹�ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		elseif (itemid="2091984") then  ''���θ����׽�Ʈ
		bufstr = bufstr & "" & vbTab & "15558147004" & vbTab & "�Ż�ǰ" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'���̹�ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		elseif (itemid="1864887") then  ''���θ����׽�Ʈ
		bufstr = bufstr & "" & vbTab & "13874181171" & vbTab & "�Ż�ǰ" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'���̹�ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		elseif (itemid="2117554") then  ''���θ����׽�Ʈ 20190425->0000000000 ���κ����غ� // ī�װ� ����1depth
		bufstr = bufstr & "50000004" & vbTab & "0000000000" & vbTab & "�Ż�ǰ" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'���̹�ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		else
		bufstr = bufstr & "" & vbTab & "" & vbTab & "�Ż�ǰ" & vbTab & importFlagYN & vbTab & "" & vbTab & isMake & vbTab		'���̹�ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		end if
		bufstr = bufstr & "" & vbTab & adultType & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab			 	'�ǸŹ�ı��� | �̼����ڱ��źҰ���ǰ���� | ��ǰ���� | ���ڵ� | ��ǰ�ڵ� | �𵨸�
		bufstr = bufstr & Replace(Replace(arrList(14,intLoop),"&nbsp;",""), vbTab, "") & vbTab & Replace(Replace(arrList(6,intLoop),"&nbsp;",""), vbTab, "") & vbTab & "" & vbTab	'�귣�� | ������ | ������
		''2021-04-01 ������ TEST
		If itemid = "1780638" Then
			bufstr = bufstr & "����ī��^120800" & vbTab		 'ī���/ī�����ΰ���
		Else
			bufstr = bufstr & "" & vbTab					 'ī���/ī�����ΰ���
		End If
		bufstr = bufstr & evtText & vbTab																			'�̺�Ʈ

		If (arrList(28,intLoop) > 0) THEN
			bufstr = bufstr & nvcpnVal & vbTab																		'�Ϲ�/��������
		ElseIf (arrList(22,intLoop) <> "") THEN
			bufstr = bufstr & Replace(arrList(22,intLoop),"&nbsp;","") & vbTab
		Else
			bufstr = bufstr & "" & vbTab
		End if

		bufstr = bufstr & isCouponDown & vbTab																		'�����ٿ�ε��ʿ俩��
		bufstr = bufstr & "" & vbTab & arrList(11,intLoop) & vbTab & "" & vbTab & "" & vbTab						'ī�幫�����Һ����� | ����Ʈ | ������ġ������ | ������Ī�ڵ�
		bufstr = bufstr & "" & vbTab	'�˻��±�..Ȯ���ʿ�
		bufstr = bufstr & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & arrList(15,intLoop) & vbTab			'�׷�ID | ���޻��ǰID | �ڵ��ǰID | �ּұ��ż��� | ��ǰ�� ����
		bufstr = bufstr & deliv & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab							'��۷� | �����ۺ񿩺� | �����ۺ񳻿� | ��ǰ�Ӽ� | ���ſɼ�
		bufstr = bufstr & "" & vbTab & "" & vbTab																	'����ID | ���̿����
		IF (isIsChangedEP) then
			bufstr = bufstr & "" & vbTab & arrList(21,intLoop) & vbTab & arrList(10,intLoop)						'���� | I,U,D | ��ǰ���������ð�
		Else
			bufstr = bufstr & ""	'����
		End If
		tFile.WriteLine bufstr
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''�ۼ��ð� üũ
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_NewCH_ST')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_NewDY_ST')"
    dbCTget.execute sqlStr
end if


if (IsChangedEP) then
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataCount](1)"
else
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataCount]"
end if
dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close

'response.write FTotCnt&"<br>"

Dim i, ArrRows, bufstr1
Dim iLastItemid : iLastItemid=9999999

IF FTotCnt > 0 THEN
    FTotPage = CLNG(FTotCnt/PageSize)
    IF FTotPage<>(FTotCnt/PageSize) THEn FTotPage=FTotPage+1
    IF (FTotPage>MaxPage) THEn
		FTotPage=MaxPage
		FTotCnt=MaxPage*PageSize
	ENd IF

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

	If (IsChangedEP) Then
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"& vbTab &"class"& vbTab &"update_time"
	Else
		bufstr1 = "id"& vbTab &"title"& vbTab &"price_pc"& vbTab &"price_mobile"& vbTab &"normal_price"& vbTab &"link"& vbTab &"mobile_link"& vbTab &"image_link"& vbTab &"add_image_link"& vbTab &"category_name1"& vbTab &"category_name2"& vbTab &"category_name3"& vbTab &"category_name4"& vbTab &"naver_category"& vbTab &"naver_product_id"& vbTab &"condition"& vbTab &"import_flag"& vbTab &"parallel_import"& vbTab &"order_made"& vbTab &"product_flag"& vbTab &"adult"& vbTab &"goods_type"& vbTab &"barcode"& vbTab &"manufacture_define_number"& vbTab &"model_number"& vbTab &"brand"& vbTab &"maker"& vbTab &"origin"& vbTab &"card_event"& vbTab &"event_words"& vbTab &"coupon"& vbTab &"partner_coupon_download"& vbTab &"interest_free_event"& vbTab &"point"& vbTab &"installation_costs"& vbTab &"pre_match_code"& vbTab &"search_tag"& vbTab &"group_id"& vbTab &"vendor_id"& vbTab &"coordi_id"& vbTab &"minimum_purchase_quantity"& vbTab &"review_count"& vbTab &"shipping"& vbTab &"delivery_grade"& vbTab &"delivery_detail"& vbTab &"attribute"& vbTab &"option_detail"& vbTab &"seller_id"& vbTab &"age_group"& vbTab &"gender"
	End If
	tFile.WriteLine bufstr1

    For i=0 to FTotPage-1
        ArrRows = ""
        if (IsChangedEP) then
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
        else
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
        end if
		dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeNaverFile(tFile,ArrRows, IsChangedEP, iLastItemid)
        end if

        ''�ۼ��ð� üũ
        IF(IsChangedEP) then
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('nvshop_NewCH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        else
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('nvshop_NewDY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        end if
    NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''�ۼ��ð� üũ
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_NewCH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_NewDY_ED')"
    dbCTget.execute sqlStr
end if

'2013-12-10 15:40 ������ �߰� TEMP������ ���� ���Ϸ� ����
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"�� ���� ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->