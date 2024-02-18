<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/shopify/shopifyItemcls.asp"-->
<!-- #include virtual="/outmall/shopify/incshopifyFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, itemoption, action, oshopify, failCnt, chgSellYn, arrRows, skipItem, tGmarketGoodno, tLimityn, getMustprice
Dim iErrStr, strParam, mustPrice, displayDate, ret1, strSql, SumErrStr, SumOKStr, iitemname, isItemIdChk
Dim i, newItemid, newItemname, strTmpGoodNo, quantity, strSKUGoodNo, maylimitEa
Dim failCnt2
Dim shopifyProductId
itemid			= requestCheckVar(request("itemid"),10)
'itemoption		= requestCheckVar(Split(request("itemid"), "_")(1),4)

If itemoption = "" Then
	itemoption = "0000"
End If

newItemid		= itemid&"_"&itemoption
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
failCnt2		= 0

Select Case action
	Case "SubCategory"	isItemIdChk = "N"
	Case Else			isItemIdChk = "Y"
End Select

If isItemIdChk = "Y" Then
	If itemid="" or itemid="0" Then
		'response.write "<script>alert('��ǰ��ȣ�� �����ϴ�.')</script>"
		'response.end
	Else
		'�������·� ��ȯ
		itemid = CLng(getNumeric(itemid))
	End If
End If

'######################################################## shopify API ########################################################
If action = "REG" Then
	SET oshopify = new Cshopify
		oshopify.FRectItemID		= itemid
		
rw "itemid:"&itemid	
		oshopify.getshopifyNotRegOneItem
		If (oshopify.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �����ϴ�."
		Else
			''newItemname = oshopify.fnshopifyItemname(itemid, itemoption, oshopify.FOneitem.FChgItemName)
			
			'' ���� ���� ���۽õ�.
			strSql = ""
			strSql = strSql & " IF NOT EXISTS(SELECT 1 FROM db_etcmall.[dbo].[tbl_shopify_regItem] WHERE itemid="&itemid&" )"
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_shopify_regItem] "
			strSql = strSql & " (itemid, regdate, reguserid, shopifystatCD, regitemname)"
			strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(newItemname)&"')"
			strSql = strSql & " END "
			strSql = strSql & " ELSE "
			strSql = strSql & " BEGIN"& VbCRLF
			strSql = strSql & " Update R"
			strSql = strSql & " set shopifystatCD='1'"
			strSql = strSql & " from db_etcmall.[dbo].[tbl_shopify_regItem] R"
			strSql = strSql & " where itemid="&itemid& VbCRLF
			strSql = strSql & " and isNULL(shopifystatCD,0)<1"
			strSql = strSql & " END "
			dbget.Execute strSql

			strParam = oshopify.FOneItem.getshopifyItemRegJSON() 
            		
			Call fnshopifyItemReg(itemid,  strParam, oshopify.FOneItem.FbasicImage, oshopify.FOneItem.FOrgprice, oshopify.FOneItem.FWonprice, oshopify.FOneItem.FMultiplerate, oshopify.FOneItem.FExchangeRate, quantity, iErrStr)
			
			'' collection ���
			if (LEFT(iErrStr, 2) = "OK") then
			    shopifyProductId  = getShopifyGoodNo(itemid)
			    Call fnShopifyCheckNRegCustomCollect(itemid, shopifyProductId)
			end if
		End If
		
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
		else
		    Call SugiQueLogInsert("shopify", action, itemid, "OK", "OK||"&itemid&"||"&iErrStr, session("ssBctID"))
		End If
	SET oshopify = nothing
ElseIf action = "EDIT" Then
	SET oshopify = new Cshopify
		oshopify.FRectItemID		= itemid
		oshopify.getshopifyEditOneItem
		
		if (oshopify.FResultCount<1) then  '' ���ǿ� ���� �ʴ� ��ǰ�� �������� ����.
		    shopifyProductId  = getShopifyGoodNo(itemid)
		    Call fnshopifyForceSlodoutProcess(itemid, shopifyProductId, iErrStr)  
		    
    	Else
    	    if (oshopify.FOneItem.FSellyn="N") then '' ǰ���� ��� �������� ����. �Ͻ�ǰ���� ����.
    	        Call fnshopifyForceSlodoutProcess(itemid, oshopify.FOneItem.FshopifyGoodNo, iErrStr) 
    	    else
        		strParam = oshopify.FOneItem.getshopifyEditJson()
''rw strParam
''response.end
        		Call fnshopifyItemEdit(itemid, oshopify.FOneItem.FshopifyGoodNo, strParam, oshopify.FOneItem.FbasicImage, oshopify.FOneItem.FOrgprice, oshopify.FOneItem.FWonprice, oshopify.FOneItem.FMultiplerate, oshopify.FOneItem.FExchangeRate, quantity, iErrStr)
        	end if
    	End If
	
    	SET oshopify = Nothing
    	
    	If LEFT(iErrStr, 2) <> "OK" Then
    		CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
    	else
		    Call SugiQueLogInsert("shopify", action, itemid, "OK", "OK||"&itemid&"||"&iErrStr, session("ssBctID"))
    	End If
    	''Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
    	
ElseIf action = "CHKSTAT" Then
	shopifyProductId = getShopifyGoodNo(itemid)
	If shopifyProductId = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||������� ���� ��ǰ�Դϴ�."
	Else
		Call fnGetShopifyGoodInfo(itemid, shopifyProductId, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
	    CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
	else
		Call SugiQueLogInsert("shopify", action, itemid, "OK", "OK||"&itemid&"||"&iErrStr, session("ssBctID"))
	End If
	    	
ElseIf action = "PRICE" Then
	SET oshopify = new Cshopify
		oshopify.FRectItemID		= itemid
		oshopify.FRectitemOption	= itemoption
		oshopify.getshopifyPriceOneItem
		If (oshopify.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||���� ���� ������ ��ǰ�� �����ϴ�."
		Else
'			strParam = ""
'			strParam = oshopify.FOneItem.getshopifyPriceBySkuNoJSON()
'			Call fnshopifyItemPriceBySkuNo(itemid, itemoption, strParam, oshopify.FOneItem.FOrgprice, oshopify.FOneItem.FWonprice, oshopify.FOneItem.FMultiplerate, oshopify.FOneItem.FExchangeRate, iErrStr)
			strParam = ""
			strParam = oshopify.FOneItem.getshopifyPriceJSON()
			Call fnshopifyItemPrice(itemid, itemoption, strParam, oshopify.FOneItem.FOrgprice, oshopify.FOneItem.FWonprice, oshopify.FOneItem.FMultiplerate, oshopify.FOneItem.FExchangeRate, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
		End If
		Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oshopify = nothing
ElseIf action = "QTY" Then
	strSKUGoodNo = getSKUshopifyGoodNo2(itemid, itemoption, quantity)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||������� ���� ��ǰ�Դϴ�."
	Else
		strParam = ""
		strParam = fnshopifyQuantityEditJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
		Call fnshopifyEditQuantity(itemid, itemoption, maylimitEa, strParam, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
	End If
	Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDITQTY" Then
	strSKUGoodNo = getSKUshopifyGoodNo2(itemid, itemoption, quantity)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||������� ���� ��ǰ�Դϴ�."
	Else
		strParam = ""
		strParam = fnshopifyQuantitySearchJSON(strSKUGoodNo)
		Call fnshopifySKUGoodNo(itemid, itemoption, strParam, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If

		If failCnt = 0 Then
			strParam = ""
			strParam = fnshopifyQuantityEditJSON(itemid, itemoption, quantity, maylimitEa, strSKUGoodNo)
			Call fnshopifyEditQuantity(itemid, itemoption, maylimitEa, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
'OK||1867487||0000||[CHKQTY]����
		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||"&itemoption&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||"&itemoption&"||", "")
			CALL Fn_AcctFailTouchEtc("shopify", itemid, SumErrStr)
			Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, "ERR", "ERR||"&itemid&"||"&itemoption&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&itemoption&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||"&itemoption&"||", "")
			Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, "OK", "OK||"&itemid&"||"&itemoption&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&itemoption&"||"&SumOKStr
		End If
	End If

ElseIf action = "RCVCOLLECTIONS" Then
    Call fnGetShopifySmartCollections(iErrStr)
    response.write iErrStr
    Call fnGetShopifyCustomCollections(iErrStr)
    response.write iErrStr
    Call fnGetShopifyCollectItems("",iErrStr)
    response.write iErrStr
    dbget.close(): response.end
ElseIf action = "CHKQUANTITY" Then
	strSKUGoodNo = getSKUshopifyGoodNo(itemid, itemoption)
	If strSKUGoodNo = "" Then
		iErrStr = "ERR||"&itemid&"||"&itemoption&"||������� ���� ��ǰ�Դϴ�."
	Else
		strParam = ""
		strParam = fnshopifyQuantitySearchJSON(strSKUGoodNo)
		Call fnshopifySKUGoodNo(itemid, itemoption, strParam, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouchEtc("shopify", itemid, iErrStr)
	End If
	Call SugiQueLogInsertByOption("shopify", action, itemid, itemoption, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "SubCategory" Then
	Call fnshopifySubCategory()
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 2000);" & vbCrLf &_
					"</script>"
End If
'###################################################### shopify API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->