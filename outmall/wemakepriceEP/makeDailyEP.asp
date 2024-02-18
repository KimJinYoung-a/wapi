<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000  ''�ʴ���
'��ǰEP�� 78�� DB�� �ٶ󺸰�, �Ǹ�EP�� 77��DB�� �ٶ󺻴�
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' ������ ���� Make / �Ϻ�
Const MaxPage   = 200   ''maxpage ���� 40->50���� 2013-12-13����, 50->60���� 2014-09-23 ������ ����, 60->70���� 2014-10-08 ���� ,70->100 ���� 2016-06-29
Const PageSize = 5000  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/wemakepriceEP/") + "\"
Dim FileName: FileName = "wemakePriceDailyEP_temp.txt"
Dim newFileName: newFileName = "wemakePriceDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "wemakePriceDailyChangedEP_temp.txt"
	newFileName = "wemakePriceDailyChangedEP.txt"
End If

Function WriteMakeWeMakePriceFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr, isMake
    Dim itemid,deliverytype, deliv, dispCash
    Dim ArrCateNM, ArrCateCD, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname, evtText, isCouponDown, nvcpnVal
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
'���ϴ� ����ī�װ�
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''��ۺ� /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 �߰�
		    ArrCateNM		= ""
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= ""
    		CntCD			= Split(ArrCateCD,",")
		Else
    		ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
    		CntNM			= Split(ArrCateNM,",")
    		ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
    		CntCD			= Split(ArrCateCD,",")
        End If
        itemname		= "[�ٹ�����]"&arrList(2,intLoop)	'2017-12-22 11:40 ������..��ǰ�� �տ� �ٹ����� �߰�
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		If (deliverytype = "7") Then deliv=-1
		If arrList(25,intLoop) = "06" OR arrList(25,intLoop) = "16" Then
			isMake = "Y"
		Else
			isMake = "N"
		End If

'		If arrList(26,intLoop) > 0 Then					'������ �������� ������...
'			dispCash	= CLNG(arrList(26,intLoop))
'			evtText		= "������������ 5% �߰����Ρ�"
'			isCouponDown= "Y"
'			nvcpnVal	= "^5"
'		Else
			dispCash	= CLNG(arrList(3,intLoop))
'			If (Now() > #13/10/2017 00:00:00# AND Now() < #25/10/2017 20:59:59#) Then
'				evtText		= "�ٹ����� 16�ֳ� ������! �ִ� 30% ��������"
'			Else
				evtText		= "�� ���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"
'			End If
			isCouponDown= ""
			nvcpnVal	= ""
'		End If

		bufstr = itemid & vbTab & Replace(itemname, vbTab, "") & vbTab & dispCash & vbTab & dispCash & vbTab  		'��ǰ�ڵ� | ��ǰ�� | pc�ǸŰ��� | ����� �ǸŰ���
		bufstr = bufstr & CLNG(arrList(24,intLoop)) & vbTab & "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=wmprc" & vbTab	'���� | ��ǰURL
		bufstr = bufstr & "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=wmprc" & vbTab									'��ǰ�����URL
		bufstr = bufstr & "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop) & vbTab & "" & vbTab	'�̹���URL | �߰� �̹���URL

		For lp = 1 to Ubound(CntNM) + 1
			If lp>4 Then Exit For
			bufstr = bufstr & Replace(CntNM(lp-1),"&nbsp;","") & vbTab																						'���޻� ī�װ���(��/��/��/��)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr & "" & vbTab
			Next
		End If

		bufstr = bufstr & "" & vbTab & "" & vbTab & "�Ż�ǰ" & vbTab & "" & vbTab & "" & vbTab & isMake & vbTab		'������ī�װ� | ���ݺ� ������ID | ��ǰ���� | �ؿܱ��Ŵ��࿩�� | ������Կ��� | �ֹ����ۻ�ǰ����
		bufstr = bufstr & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab			 	'�ǸŹ�ı��� | �̼����ڱ��źҰ���ǰ���� | ��ǰ���� | ���ڵ� | ��ǰ�ڵ� | �𵨸�
		bufstr = bufstr & Replace(Replace(arrList(14,intLoop),"&nbsp;",""), vbTab, "") & vbTab & Replace(Replace(arrList(6,intLoop),"&nbsp;",""), vbTab, "") & vbTab & "" & vbTab & "" & vbTab		'�귣�� | ������ | ������ | ī���/ī�����ΰ���
		bufstr = bufstr & evtText & vbTab																			'�̺�Ʈ

'		If (arrList(26,intLoop) > 0) THEN
'			bufstr = bufstr & nvcpnVal & vbTab																		'�Ϲ�/��������
'		ElseIf (arrList(22,intLoop) <> "") THEN
		If (arrList(22,intLoop) <> "") THEN
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
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_CH_ST')"
    dbCTget.execute sqlStr
else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_DY_ST')"
    dbCTget.execute sqlStr
end if


if (IsChangedEP) then
    sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPDataCount](1)"
else
    sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPDataCount]"
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
	end if

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
            sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
        else
            sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
        end if
		dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeWeMakePriceFile(tFile,ArrRows, IsChangedEP, iLastItemid)
        end if

        ''�ۼ��ð� üũ
        IF(IsChangedEP) then
            sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('wemakePrice_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        else
            sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('wemakePrice_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        end if
    NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''�ۼ��ð� üũ
IF(IsChangedEP) then
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_CH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('wemakePrice_DY_ED')"
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