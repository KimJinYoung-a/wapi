<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000 ''�ʴ���
'��ǰEP�� 78�� DB�� �ٶ󺸰�, �Ǹ�EP�� 77��DB�� �ٶ󺻴�
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 1   ''maxpage 100..2015-05-07 ������ ����
Const PageSize = 100  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/wemakepriceEP/") + "\"
Dim FileName: FileName = "HotDeal_temp.xml"
Dim newFileName: newFileName = "HotDeal.xml"
Dim fso, tFile

Function WriteMakeWeMakePriceFile(tFile, arrList, byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr
    Dim itemid, deliv, ArrCateNM, CntNM
    Dim lp, lp2, limitSu, itemdivCost, freeDlvYn
    Dim tmpLastDeptNM, itemname, couponCash, couponPer
	Dim isDealItem
    iRow = UBound(arrList,2)

		bufstr = "<?xml version=""1.0"" encoding=""euc-kr"" ?>"&VbCRLF
		bufstr = bufstr&"<products>"&VbCRLF													'*## ��ü ��ǰ ������ ����/�� | ��ü ��ǰ ������ ���ۺ��� ���� <products> ~ </products>�� �����Ͽ� ����
		tFile.WriteLine bufstr
	For intLoop=0 to iRow
		itemid			= arrList(1,intLoop)
		itemname		= "[�ٹ�����]"&arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		ArrCateNM		= Split(arrList(12,intLoop),"||")(0)
		CntNM			= Split(ArrCateNM,",")

		itemdivCost		= arrList(8,intLoop)
		If itemdivCost > 0 Then
			freeDlvYn = "Y"
		Else
			freeDlvYn = "N"
		End If

		If arrList(9,intLoop) = "Y" Then
			limitSu = arrList(10,intLoop) - arrList(11,intLoop)
		Else
			limitSu = "999999"
		End If

		isDealItem = (arrList(16,intLoop)="21")

		bufstr = ""
		bufstr = bufstr&"<product>"&VbCRLF
		bufstr = bufstr&"<mallPid>"&itemid&"</mallPid>"&VbCRLF									'*�� ��ǰ ���̵�(��ǰ��ȣ)
		bufstr = bufstr&"<poplrDgr>"&arrList(0,intLoop)&"</poplrDgr>"&VbCRLF					'*�α⵵
		bufstr = bufstr&"<prodName><![CDATA["&itemname&"]]></prodName>"&VbCRLF					'*��ǰ��
		if (isDealItem) then
			bufstr = bufstr&"<prodUrl><![CDATA[http://www.10x10.co.kr/deal/deal.asp?itemid="&itemid&"&rdsite=wmprchot]]></prodUrl>"&VbCRLF				'*pc ���� ������
			bufstr = bufstr&"<mblProdUrl><![CDATA[http://m.10x10.co.kr/deal/deal.asp?itemid="&itemid&"&rdsite=wmprchot]]></mblProdUrl>"&VbCRLF		'*����� ����������
			bufstr = bufstr&"<prodImgUrl1><![CDATA[http://webimage.10x10.co.kr/image/basic/" & arrList(7,intLoop)&"]]></prodImgUrl1>"&VbCRLF	'*��ǥ �̹���URL(���簢��)
		else
			bufstr = bufstr&"<prodUrl><![CDATA[http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=wmprchot]]></prodUrl>"&VbCRLF				'*pc ���� ������
			bufstr = bufstr&"<mblProdUrl><![CDATA[http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=wmprchot]]></mblProdUrl>"&VbCRLF		'*����� ����������
			bufstr = bufstr&"<prodImgUrl1><![CDATA[http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(7,intLoop)&"]]></prodImgUrl1>"&VbCRLF	'*��ǥ �̹���URL(���簢��)
		end if
		bufstr = bufstr&"<nomlPrc>"&CLng(arrList(3,intLoop))&"</nomlPrc>"&VbCRLF				'*���� ����
		bufstr = bufstr&"<mblDcPrc>"&Clng(arrList(5,intLoop))&"</mblDcPrc>"&VbCRLF				'*����� ���ΰ���
		bufstr = bufstr&"<mblDcRt>"&CInt(arrList(6,intLoop))&"</mblDcRt>"&VbCRLF				'����� ������
		bufstr = bufstr&"<freeDlvYn>"&freeDlvYn&"</freeDlvYn>"&VbCRLF							'*������ ����
		bufstr = bufstr&"<saleCnt>"&arrList(15,intLoop)&"</saleCnt>"&VbCRLF						'*���� ����(�� ���� �� ����)
		For lp=1 to Ubound(CntNM)+1
			If lp>5 Then Exit For
			bufstr = bufstr&"<catNm"&lp&"><![CDATA["&Replace(CntNM(lp-1),"&nbsp;","")&"]]></catNm"&lp&">"&VbCRLF	'*���޻� ī�װ���(��/��/��/���з� ī�װ�)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr&"<catNm"&lp&"></catNm"&lp&">"&VbCRLF
			Next
		End If
		bufstr = bufstr&"<modelNm></modelNm>"&VbCRLF											'�𵨸�
		bufstr = bufstr&"<brandNm><![CDATA["&arrList(13,intLoop)&"]]></brandNm>"&VbCRLF			'�귣���
		bufstr = bufstr&"</product>"&VbCRLF
		tFile.WriteLine bufstr
		bufstr = ""
		iLastItemid = itemid
	Next
		'�� �κ��� �ݺ��Ǿ�� �� / ��
		bufstr = bufstr&"</products>"
		tFile.WriteLine bufstr
		bufstr = ""
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('wemakePrice_HOT_ST')"
dbCTget.execute sqlStr

Dim i, ArrRows
Dim iLastItemid : iLastItemid=9999999

Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

		ArrRows = ""
		sqlStr ="[db_outmall].[dbo].[usp_Ten_Outmall_Wemakeprice_EPData]("&i+1&","&PageSize&",2,"&iLastItemid&")"
		dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
	    rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	    IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	    	ArrRows = rsCTget.getRows()
	    END IF
	    rsCTget.close

	    if isArray(ArrRows) then
	        CALL WriteMakeWeMakePriceFile(tFile, ArrRows, iLastItemid)
	    end if

		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
		sqlStr = sqlStr + " (ref) values('wemakePrice_HOT"&(i+1)*PageSize&"_"&iLastItemid&"')"
		dbCTget.execute sqlStr

	    tFile.Close
	Set tFile = Nothing
Set fso = Nothing

sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('wemakePrice_HOT_ED')"
dbCTget.execute sqlStr

'2013-12-10 15:40 ������ �߰� TEMP������ ���� ���Ϸ� ����
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write "100�� ���� ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->