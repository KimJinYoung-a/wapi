<%
'############################################## ���� �����ϴ� API �Լ� ���� ���� ############################################
'' Smart Collection
Public Function fnGetShopifySmartCollections(byRef iErrStr)
    Dim objJSON, istrParam, iRbody, strObj, oneColection
    Dim ColCnt, i, j
    Dim ColectionId, ColectionTitle, ColectionHandle, ColectionUpdated_at, ColectionPublished_at, ColectionRules
    Dim ColectionRule1_column,ColectionRule1_relation, ColectionRule1_condition
    Dim ColectionRule2_column,ColectionRule2_relation, ColectionRule2_condition
    
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objJSON.Open "GET", shopifyAPIURL & "/admin/smart_collections.json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send()
        If objJSON.Status = "200" Then  
            iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			''response.write iRbody
			Set strObj = JSON.parse(iRbody)
			
			if (strObj.smart_collections="") then
			    iErrStr = "ERR||Collections��  �����ϴ�..[ERR-CHKSTAT]"
		    else
		        ColCnt = strObj.smart_collections.length
		        for i=0 to ColCnt-1
		        
    			    SET oneColection = strObj.smart_collections.get(i)
    				ColectionId = oneColection.id
    				ColectionTitle = oneColection.title
    				ColectionHandle = oneColection.handle
    				ColectionUpdated_at = replace(LEFT(oneColection.updated_at,19),"T"," ")
    				ColectionPublished_at = replace(LEFT(oneColection.published_at,19),"T"," ")
    				SET ColectionRules = oneColection.rules
    				
    				for j=0 to ColectionRules.length-1
    				    if (j=0) then
        				    ColectionRule1_column    = ColectionRules.get(0).column
        				    ColectionRule1_relation  = ColectionRules.get(0).relation
        				    ColectionRule1_condition = ColectionRules.get(0).condition
        				elseif (j=1) then
        				    ColectionRule2_column    = ColectionRules.get(1).column
        				    ColectionRule2_relation  = ColectionRules.get(1).relation
        				    ColectionRule2_condition = ColectionRules.get(1).condition
        				else
        				    '' skip
        				end if
    			    next
    				SET ColectionRules = Nothing
    				
    				''rw ColectionId&"-"&ColectionTitle
    				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateCollections] "&ColectionId&",'"&html2db(ColectionTitle)&"','"&ColectionHandle&"','"&ColectionUpdated_at&"','"&ColectionPublished_at&"','"&ColectionRule1_column&"','"&ColectionRule1_relation&"','"&ColectionRule1_condition&"','"&ColectionRule2_column&"','"&ColectionRule2_relation&"','"&ColectionRule2_condition&"'"
        
                    ''rw strSql
        			dbget.execute strSql
        				
    				SET oneColection = Nothing
    			Next
			end if
			
            Set strObj = Nothing
		Else
			iErrStr = "ERR||"&objJSON.Status&"||shopify Smart Collection ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
		
	Set objJSON= nothing
end function

'' Custom Collection
Public Function fnGetShopifyCustomCollections(byRef iErrStr)
    Dim objJSON, istrParam, iRbody, strObj, oneColection
    Dim ColCnt, i, j
    Dim ColectionId, ColectionTitle, ColectionHandle, ColectionUpdated_at, ColectionPublished_at, ColectionRules
    Dim ColectionRule1_column,ColectionRule1_relation, ColectionRule1_condition
    Dim ColectionRule2_column,ColectionRule2_relation, ColectionRule2_condition
    
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objJSON.Open "GET", shopifyAPIURL & "/admin/custom_collections.json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send()
        If objJSON.Status = "200" Then  
            iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			''response.write iRbody
			Set strObj = JSON.parse(iRbody)
			
			if (strObj.custom_collections="") then
			    iErrStr = "ERR||Collections��  �����ϴ�..[ERR-CHKSTAT]"
		    else
		        ColCnt = strObj.custom_collections.length
		        for i=0 to ColCnt-1
		        
    			    SET oneColection = strObj.custom_collections.get(i)
    				ColectionId = oneColection.id
    				ColectionTitle = oneColection.title
    				ColectionHandle = oneColection.handle
    				ColectionUpdated_at = replace(LEFT(oneColection.updated_at,19),"T"," ")
    				ColectionPublished_at = replace(LEFT(oneColection.published_at,19),"T"," ")
    				
    				''rw ColectionId&"-"&ColectionTitle
    				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateCollections] "&ColectionId&",'"&html2db(ColectionTitle)&"','"&ColectionHandle&"','"&ColectionUpdated_at&"','"&ColectionPublished_at&"','"&ColectionRule1_column&"','"&ColectionRule1_relation&"','"&ColectionRule1_condition&"','"&ColectionRule2_column&"','"&ColectionRule2_relation&"','"&ColectionRule2_condition&"'"
        
                    ''rw strSql
        			dbget.execute strSql
        				
    				SET oneColection = Nothing
    			Next
			end if
			
            Set strObj = Nothing
		Else
			iErrStr = "ERR||"&objJSON.Status&"||shopify Custom Collection ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
		
	Set objJSON= nothing
end function

'' collection ��ǰ���
public function fnGetShopifyGetOneCollectItems(icollectionid, byRef iErrStr)
    Dim objJSON, istrParam, iRbody, strObj, oneColection
    Dim ColCnt, i, j
    Dim ColectId, ColectionId, ColectProduct_id, ColectFeatured, ColectCreated_at, ColectUpdated_at
    Dim ColectPosition
    Dim ColectSort_value
    
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objJSON.Open "GET", shopifyAPIURL & "/admin/collects.json?collection_id="&icollectionid , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send()
        If objJSON.Status = "200" Then  
            iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			''response.write iRbody
			Set strObj = JSON.parse(iRbody)
			
			if (strObj.collects="") then
			    iErrStr = "ERR||Collections��  �����ϴ�..[ERR-CHKSTAT]"
		    else
		        ColCnt = strObj.collects.length
		        for i=0 to ColCnt-1
		        
    			    SET oneColection = strObj.collects.get(i)
    			    ColectId = oneColection.id
    				ColectionId = oneColection.collection_id
    				ColectProduct_id = oneColection.product_id
    				ColectFeatured = oneColection.featured
    				ColectCreated_at = replace(LEFT(oneColection.created_at,19),"T"," ")
    				ColectUpdated_at = replace(LEFT(oneColection.updated_at,19),"T"," ")
    				ColectPosition = oneColection.position
    				ColectSort_value = oneColection.sort_value
    				
    				''rw ColectId&"-"&ColectionId&"-"&ColectProduct_id&":"&ColectPosition&":"&ColectSort_value
    				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateCollectionItem] "&ColectId&",'"&ColectionId&"','"&ColectProduct_id&"','"&ColectFeatured&"','"&ColectCreated_at&"','"&ColectUpdated_at&"',"&ColectPosition&""
        
                    ''rw strSql
        			dbget.execute strSql
        				
    				SET oneColection = Nothing
    			Next
			end if
			
            Set strObj = Nothing
		Else
			iErrStr = "ERR||"&objJSON.Status&"||shopify Custom Collection ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
		
	Set objJSON= nothing
end function

'' CollectItems 
Public Function fnGetShopifyCollectItems(icollectionid, byRef iErrStr)
    dim sqlStr : sqlStr = "exec db_etcmall.[dbo].[usp_TEN_Shopify_GetCollectionList] '"&icollectionid&"'"
    dim ArrRows , acollectionid, lp
    dim reError, iErr
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		ArrRows = rsget.getRows()
	rsget.Close
	
	if isArray(ArrRows) then
	    For lp = 0 To Ubound(ArrRows, 2)
	        acollectionid = ArrRows(0,lp)
	        ''rw acollectionid
            Call fnGetShopifyGetOneCollectItems(icollectionid, iErr)
            reError = reError & iErr
        Next
    end if
    iErrStr = reError
end function

public Function fnShopifyRegCustomCollection(collectionTitle,collectionType, customcollectiontp_val)
    Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody, strObj
    Dim istrParam
    Dim oneColection
    Dim ColectionId, retColectionTitle, ColectionHandle, ColectionUpdated_at, ColectionPublished_at
    Dim ColectionRule1_column,ColectionRule1_relation,ColectionRule1_condition,ColectionRule2_column,ColectionRule2_relation,ColectionRule2_condition
    
    fnShopifyRegCustomCollection = ""
    istrParam = ""
	istrParam = istrParam & "{"
    istrParam = istrParam & " ""custom_collection"": {"
    istrParam = istrParam & " ""title"": """&collectionTitle&""""
    istrParam = istrParam & " }"
    istrParam = istrParam & "}"
        
        
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/admin/custom_collections.json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send(istrParam)
 
		If objJSON.Status = "201" Then  ''HTTP/1.1 201 Created
		    iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			response.write iRbody
			Set strObj = JSON.parse(iRbody)
			if (strObj.custom_collection="") then
			    rw "ERR||Collections��  �����ϴ�..[ERR-CHKSTAT]"
			    exit function
		    else
			    SET oneColection = strObj.custom_collection
				ColectionId = oneColection.id
				retColectionTitle = oneColection.title
				ColectionHandle = oneColection.handle
				ColectionUpdated_at = replace(LEFT(oneColection.updated_at,19),"T"," ")
				ColectionPublished_at = replace(LEFT(oneColection.published_at,19),"T"," ")
				
				rw ColectionId&"-"&retColectionTitle
				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateCollections] "&ColectionId&",'"&html2db(retColectionTitle)&"','"&ColectionHandle&"','"&ColectionUpdated_at&"','"&ColectionPublished_at&"','"&ColectionRule1_column&"','"&ColectionRule1_relation&"','"&ColectionRule1_condition&"','"&ColectionRule2_column&"','"&ColectionRule2_relation&"','"&ColectionRule2_condition&"',"&collectionType&",'"&customcollectiontp_val&"'"
    
                ''rw strSql
    			dbget.execute strSql
    				
				SET oneColection = Nothing
			end if
		    Set strObj = Nothing
		Else
			rw "ERR||"&collectionTitle&"||"&objJSON.Status&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objJSON = Nothing
	
	fnShopifyRegCustomCollection = ColectionId
end function

'' Custom Collect ���
Public Function fnShopifyCheckNRegCustomCollect(iitemid, productid)
    '' �ش� Collection �� ��ϵǾ� �ִ��� �˻� �� ����.
    dim icollectionId 
    icollectionId = chkMakeCollection(iitemid,10001)  '' �귣��
    if (icollectionId<>"") then
        Call fnShopifyRegCustomCollect(iitemid,productid,icollectionId)
    end if
    
    icollectionId = chkMakeCollection(iitemid,10002) '' Cate2depth
    if (icollectionId<>"") then
        Call fnShopifyRegCustomCollect(iitemid,productid,icollectionId)
    end if

end function

Public Function fnShopifyRegCustomCollect(iitemid,productid,collectionId)
    Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody, strObj
    Dim istrParam
    Dim oneColect
    Dim ColectId, retColectionId, ColectProduct_id, ColectFeatured, ColectCreated_at, ColectUpdated_at
    Dim ColectPosition
    Dim ColectSort_value
    Dim idbcollectId
    
    strSql = "exec db_etcmall.dbo.usp_TEN_Shopify_CheckCollectItemExists '"&collectionId&"','"&productid&"'" 
    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
    If not rsget.EOF Then
		idbcollectId = rsget("collectId")
	End If
    rsget.close
    
    if (idbcollectId<>"") then
        fnShopifyRegCustomCollect = idbcollectId
        Exit function
    end if
    
    fnShopifyRegCustomCollect = ""
    istrParam = ""
	istrParam = istrParam & "{"
    istrParam = istrParam & " ""collect"": {"
    istrParam = istrParam & " ""product_id"": """&productid&""","
    istrParam = istrParam & " ""collection_id"": """&collectionId&""""
    istrParam = istrParam & " }"
    istrParam = istrParam & "}"
        
        
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/admin/collects.json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
''rw shopifyAPIURL & "/admin/collects.json"		
		objJSON.Send(istrParam)
 
		If objJSON.Status = "201" Then  ''HTTP/1.1 201 Created
		    iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
			
			if (strObj.collect="") then
			    iErrStr = "ERR||Collections��  �����ϴ�..[ERR-CHKSTAT]"
		    else
		        
			    SET oneColect = strObj.collect
			    ColectId = oneColect.id
				retColectionId = oneColect.collection_id
				ColectProduct_id = oneColect.product_id
				ColectFeatured = oneColect.featured
				ColectCreated_at = replace(LEFT(oneColect.created_at,19),"T"," ")
				ColectUpdated_at = replace(LEFT(oneColect.updated_at,19),"T"," ")
				ColectPosition = oneColect.position
				ColectSort_value = oneColect.sort_value
				
				''rw ColectId&"-"&retColectionId&"-"&ColectProduct_id&":"&ColectPosition&":"&ColectSort_value
				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateCollectionItem] "&ColectId&",'"&retColectionId&"','"&ColectProduct_id&"','"&ColectFeatured&"','"&ColectCreated_at&"','"&ColectUpdated_at&"',"&ColectPosition&""
    
                ''rw strSql
    			dbget.execute strSql
    				
				SET oneColect = Nothing
			end if
			
            Set strObj = Nothing
			
		Else
		    '' �̹� ��ϵǾ� ������ Status:422�� ��ȯ�Ѵ�.
			rw "ERR||"&collectionId&"||"&objJSON.Status&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objJSON = Nothing
	
	fnShopifyRegCustomCollect = ColectId
end function

'' collection �� ������ �����Ѵ�.
public Function chkMakeCollection(iitemid, collectionType)
    ''
    Dim sqlStr ,collectionId, collectionTitle, customcollectiontp_val
    sqlStr = "exec db_etcmall.dbo.usp_TEN_Shopify_CheckCustomCollectionExists "&iitemid&","&collectionType&""  ''�귣��

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    If not rsget.EOF Then
		collectionId = rsget("collectionId")
		collectionTitle = rsget("collectionTitle")
		customcollectiontp_val = rsget("customcollectiontp_val")
	End If
    rsget.close

    if isNULL(collectionId) then
        if (collectionTitle<>"") then
            collectionId = fnShopifyRegCustomCollection(collectionTitle,collectionType,customcollectiontp_val)
        end if
    end if
    
    chkMakeCollection = collectionId
end function

'��ǰ ���
Public Function fnshopifyItemReg(iitemid, istrParam, iimageNm, iOrgprice, irateprice, iMultiplerate, iExchangeRate, iquantity, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess, productClientId
	
	Dim oneProduct
	Dim shopifyGoodNo,shopifyTitle,shopifyProduct_type,shopifyHandle,shopifyUpdated_at,shopifyPublished_at,shopifyPublished_scope,shopifyTags
	Dim shopifySKUId, shopifySKUtitle, shopifySKUprice, shopifySKUcompare_at_price, shopifySKUoption1, shopifySKUoption2, shopifySKUoption3
	Dim shopifySKUsku, shopifySKUbarcode, shopifySKUgrams, shopifySKUinventory_quantity, shopifySKUold_inventory_quantity, shopifySKUposition
	Dim shopifysellYn, shopifyStatCd
	
'	On Error Resume Next
	
		
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/admin/products.json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send(istrParam)
 
		If objJSON.Status = "201" Then  ''HTTP/1.1 201 Created
			iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			''response.write iRbody
			Set strObj = JSON.parse(iRbody)
					SET oneProduct = strObj.product
    				shopifyGoodNo = oneProduct.id
    				shopifyTitle = oneProduct.title
    				shopifyProduct_type = oneProduct.product_type
    				shopifyHandle = oneProduct.handle
    				shopifyUpdated_at = replace(LEFT(oneProduct.updated_at,19),"T"," ")
    				shopifyPublished_at = replace(LEFT(oneProduct.published_at,19),"T"," ")
    				shopifyPublished_scope = oneProduct.published_scope
    				shopifyTags = oneProduct.tags
    				
    				''rw shopifyGoodNo&"|"&shopifyTitle&"|"&shopifyProduct_type&"|"&shopifyHandle&"|"&shopifyUpdated_at&"|"&shopifyPublished_scope&"|"&shopifyTags
    				
    				'strStatus	  = strObj.status
    				For i=0 to oneProduct.variants.length-1
    					shopifySKUId = oneProduct.variants.get(i).id
    					shopifySKUtitle = oneProduct.variants.get(i).title
    					shopifySKUprice = oneProduct.variants.get(i).price
    					shopifySKUcompare_at_price = oneProduct.variants.get(i).compare_at_price
    					shopifySKUoption1 = oneProduct.variants.get(i).option1
    					shopifySKUoption2 = oneProduct.variants.get(i).option2
    					shopifySKUoption3 = oneProduct.variants.get(i).option3
    					shopifySKUsku       = oneProduct.variants.get(i).sku
    					shopifySKUbarcode   = oneProduct.variants.get(i).barcode
    					shopifySKUgrams     = oneProduct.variants.get(i).grams
    					shopifySKUinventory_quantity = oneProduct.variants.get(i).inventory_quantity
    					shopifySKUold_inventory_quantity = oneProduct.variants.get(i).old_inventory_quantity
    					shopifySKUposition = oneProduct.variants.get(i).position
    					''shopifySKUinventory_item_id = oneProduct.variants.get(i).inventory_item_id
    					''weight
    					''weight_unit
    					
    					''rw shopifySKUId&"|"&shopifySKUtitle&"|"&shopifySKUprice&"|"&shopifySKUcompare_at_price&"|"&shopifySKUoption1&"|"&shopifySKUoption2&"|"&shopifySKUoption3&"|"&shopifySKUsku&"|"&shopifySKUbarcode&"|"&shopifySKUgrams&"|"&shopifySKUgrams
    				Next
    				
    				if (shopifyPublished_at="null") then
    				    shopifysellYn = "N"
        			else
        			    shopifysellYn = "Y"
        			end if
        			
        			if (shopifyPublished_at="null") then
        			    shopifyStatCd = "3"
        			else
        			    shopifyStatCd = "7"
        		    end if

                    strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateMappingItem] "&iitemid&",'"&shopifyPublished_at&"','"&shopifyUpdated_at&"','"&shopifyGoodNo&"','"&html2db(shopifyHandle)&"',"&shopifySKUprice&","&shopifySKUcompare_at_price&",'"&shopifySellYn&"',"&shopifyStatCd&",'"&html2db(shopifyTitle)&"','"&iimageNm&"'"
    
                    ''rw strSql
    				dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||[REG]����"
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&objJSON.Status&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG]"
		End If
	Set objJSON= nothing
End Function

'' ���� ǰ��(���þ���) ó��
function fnshopifyForceSlodoutProcess(iitemid, ishopifyProductId,byRef iErrStr)  
    Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody, strObj
    Dim istrParam
    Dim oneProduct
	Dim shopifyGoodNo,shopifyTitle,shopifyProduct_type,shopifyHandle,shopifyUpdated_at,shopifyPublished_at,shopifyPublished_scope,shopifyTags
	Dim shopifySKUId, shopifySKUtitle, shopifySKUprice, shopifySKUcompare_at_price, shopifySKUoption1, shopifySKUoption2, shopifySKUoption3
	Dim shopifySKUsku, shopifySKUbarcode, shopifySKUgrams, shopifySKUinventory_quantity, shopifySKUold_inventory_quantity, shopifySKUposition
	Dim shopifysellYn, shopifyStatCd
    
    
    fnshopifyForceSlodoutProcess = false
    istrParam = ""
	istrParam = istrParam & "{"
    istrParam = istrParam & " ""product"": {"
    istrParam = istrParam & " ""id"": """&ishopifyProductId&""","
    istrParam = istrParam & " ""published"": ""false"""
    istrParam = istrParam & " }"
    istrParam = istrParam & "}"
        
        
    Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "PUT", shopifyAPIURL & "/admin/products/"&ishopifyProductId&".json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"

		objJSON.Send(istrParam)
'rw shopifyAPIURL & "/admin/products/"&ishopifyProductId&".json"
'rw istrParam
		If objJSON.Status = "200" Then  ''HTTP/1.1 200 ok
			iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
rw iRbody
			Set strObj = JSON.parse(iRbody)
					SET oneProduct = strObj.product
    				shopifyGoodNo = oneProduct.id
    				shopifyTitle = oneProduct.title
    				shopifyProduct_type = oneProduct.product_type
    				shopifyHandle = oneProduct.handle
    				shopifyUpdated_at = replace(LEFT(oneProduct.updated_at,19),"T"," ")
    				if isNULL(oneProduct.published_at) then
    				    shopifyPublished_at = "null"
    				else
    				    shopifyPublished_at = replace(LEFT(oneProduct.published_at,19),"T"," ")
    			    end if
    				shopifyPublished_scope = oneProduct.published_scope
    				shopifyTags = oneProduct.tags
    				
    				''rw shopifyGoodNo&"|"&shopifyTitle&"|"&shopifyProduct_type&"|"&shopifyHandle&"|"&shopifyUpdated_at&"|"&shopifyPublished_scope&"|"&shopifyTags
    				
    				'strStatus	  = strObj.status
    				For i=0 to oneProduct.variants.length-1
    					shopifySKUId = oneProduct.variants.get(i).id
    					shopifySKUtitle = oneProduct.variants.get(i).title
    					shopifySKUprice = oneProduct.variants.get(i).price
    					shopifySKUcompare_at_price = oneProduct.variants.get(i).compare_at_price
    					shopifySKUoption1 = oneProduct.variants.get(i).option1
    					shopifySKUoption2 = oneProduct.variants.get(i).option2
    					shopifySKUoption3 = oneProduct.variants.get(i).option3
    					shopifySKUsku       = oneProduct.variants.get(i).sku
    					shopifySKUbarcode   = oneProduct.variants.get(i).barcode
    					shopifySKUgrams     = oneProduct.variants.get(i).grams
    					shopifySKUinventory_quantity = oneProduct.variants.get(i).inventory_quantity
    					shopifySKUold_inventory_quantity = oneProduct.variants.get(i).old_inventory_quantity
    					shopifySKUposition = oneProduct.variants.get(i).position
    					''shopifySKUinventory_item_id = oneProduct.variants.get(i).inventory_item_id
    					''weight
    					''weight_unit
    					
    					''rw shopifySKUId&"|"&shopifySKUtitle&"|"&shopifySKUprice&"|"&shopifySKUcompare_at_price&"|"&shopifySKUoption1&"|"&shopifySKUoption2&"|"&shopifySKUoption3&"|"&shopifySKUsku&"|"&shopifySKUbarcode&"|"&shopifySKUgrams&"|"&shopifySKUgrams
    				Next
    				
    				if (shopifyPublished_at="null") then
    				    shopifysellYn = "N"
        			else
        			    shopifysellYn = "Y"
        			end if
        			
        			if (shopifyPublished_at="null") then
        			    shopifyStatCd = "3"
        			else
        			    shopifyStatCd = "7"
        		    end if

                    strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateMappingItem] "&iitemid&",'"&shopifyPublished_at&"','"&shopifyUpdated_at&"','"&shopifyGoodNo&"','"&html2db(shopifyHandle)&"',"&shopifySKUprice&","&shopifySKUcompare_at_price&",'"&shopifySellYn&"',"&shopifyStatCd&",'"&html2db(shopifyTitle)&"',''"
    
                    ''rw strSql
    				dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||[SOLDOUT]����"
			Set strObj = Nothing
			fnshopifyForceSlodoutProcess = true
		Else
			iErrStr = "ERR||"&iitemid&"||"&objJSON.Status&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SOLDOUT]"
		End If
	Set objJSON = Nothing
	
	
end function


function fnshopifyItemEdit(iitemid, ishopifyProductId, istrParam, iimageNm, iOrgprice, irateprice, iMultiplerate, iExchangeRate, iquantity, byRef iErrStr)
    Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess, productClientId
	
	Dim oneProduct
	Dim shopifyGoodNo,shopifyTitle,shopifyProduct_type,shopifyHandle,shopifyUpdated_at,shopifyPublished_at,shopifyPublished_scope,shopifyTags
	Dim shopifySKUId, shopifySKUtitle, shopifySKUprice, shopifySKUcompare_at_price, shopifySKUoption1, shopifySKUoption2, shopifySKUoption3
	Dim shopifySKUsku, shopifySKUbarcode, shopifySKUgrams, shopifySKUinventory_quantity, shopifySKUold_inventory_quantity, shopifySKUposition
	Dim shopifysellYn, shopifyStatCd
	Dim ttlquantity : ttlquantity = 0
'	On Error Resume Next
	
		
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "PUT", shopifyAPIURL & "/admin/products/"&ishopifyProductId&".json" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
		
		objJSON.Send(istrParam)
 
		If objJSON.Status = "200" Then  ''HTTP/1.1 200 OK
			iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			''response.write iRbody
			Set strObj = JSON.parse(iRbody)
					SET oneProduct = strObj.product
    				shopifyGoodNo = oneProduct.id
    				shopifyTitle = oneProduct.title
    				shopifyProduct_type = oneProduct.product_type
    				shopifyHandle = oneProduct.handle
    				shopifyUpdated_at = replace(LEFT(oneProduct.updated_at,19),"T"," ")
    				shopifyPublished_at = replace(LEFT(oneProduct.published_at,19),"T"," ")
    				shopifyPublished_scope = oneProduct.published_scope
    				shopifyTags = oneProduct.tags
    				
    				''rw shopifyGoodNo&"|"&shopifyTitle&"|"&shopifyProduct_type&"|"&shopifyHandle&"|"&shopifyUpdated_at&"|"&shopifyPublished_scope&"|"&shopifyTags
    				
    				'strStatus	  = strObj.status
    				For i=0 to oneProduct.variants.length-1
    					shopifySKUId = oneProduct.variants.get(i).id
    					shopifySKUtitle = oneProduct.variants.get(i).title
    					shopifySKUprice = oneProduct.variants.get(i).price
    					shopifySKUcompare_at_price = oneProduct.variants.get(i).compare_at_price
    					shopifySKUoption1 = oneProduct.variants.get(i).option1
    					shopifySKUoption2 = oneProduct.variants.get(i).option2
    					shopifySKUoption3 = oneProduct.variants.get(i).option3
    					shopifySKUsku       = oneProduct.variants.get(i).sku
    					shopifySKUbarcode   = oneProduct.variants.get(i).barcode
    					shopifySKUgrams     = oneProduct.variants.get(i).grams
    					shopifySKUinventory_quantity = oneProduct.variants.get(i).inventory_quantity
    					shopifySKUold_inventory_quantity = oneProduct.variants.get(i).old_inventory_quantity
    					shopifySKUposition = oneProduct.variants.get(i).position
    					
    					if isNumeric(shopifySKUinventory_quantity) then
    					    ttlquantity = ttlquantity+CLNG(shopifySKUinventory_quantity)
    				    end if
    					''shopifySKUinventory_item_id = oneProduct.variants.get(i).inventory_item_id
    					''weight
    					''weight_unit
    					
    					''rw shopifySKUId&"|"&shopifySKUtitle&"|"&shopifySKUprice&"|"&shopifySKUcompare_at_price&"|"&shopifySKUoption1&"|"&shopifySKUoption2&"|"&shopifySKUoption3&"|"&shopifySKUsku&"|"&shopifySKUbarcode&"|"&shopifySKUgrams&"|"&shopifySKUgrams
    				Next
    				
    				if (shopifyPublished_at="null") then
    				    shopifysellYn = "N"
        			else
        			    shopifysellYn = "Y"
        			end if
        			
        			if (shopifyPublished_at="null") then
        			    shopifyStatCd = "3"
        			else
        			    shopifyStatCd = "7"
        		    end if

                    strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateMappingItem] "&iitemid&",'"&shopifyPublished_at&"','"&shopifyUpdated_at&"','"&shopifyGoodNo&"','"&html2db(shopifyHandle)&"',"&shopifySKUprice&","&shopifySKUcompare_at_price&",'"&shopifySellYn&"',"&shopifyStatCd&",'"&html2db(shopifyTitle)&"','"&iimageNm&"',"&ttlquantity&""
    
                    ''rw strSql
    				dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||[EDIT]����"
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&objJSON.Status&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT]"
		End If
	Set objJSON= nothing
end function

'��ǰ ��ȸ ��ü. �ʱ���.
'Public Function fnshopifyItemListAll(iitemid, iitemoption, istrTmpGoodNo, byRef iErrStr)
'	Dim objJSON, iRbody, strObj, strSql, strStatus, i
'	Dim shopifyGoodNo, shopifyTitle, shopifyProduct_type, shopifyHandle, shopifyUpdated_at, shopifyPublished_at, shopifyPublished_scope, shopifyTags
'	Dim shopifySKUId, shopifySKUtitle, shopifySKUprice,shopifySKUcompare_at_price,shopifySKUoption1,shopifySKUoption2,shopifySKUoption3,shopifySKUsku,shopifySKUbarcode,shopifySKUgrams,shopifySKUinventory_quantity
'	Dim shopifySKUold_inventory_quantity,shopifySKUinventory_item_id,shopifySKUposition
'	
'	Dim oneProduct
'	On Error Resume Next
'	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
'	    objJSON.Open "GET", shopifyAPIURL & "/admin/products.json" , False ,shopifySELLERID,shopifyAPIKEY
'	 
'		objJSON.setRequestHeader "Content-Type", "application/json"
'		objJSON.Send()
'		If objJSON.Status = "200" Then
'			iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
''			response.write iRbody
'			
'			Set strObj = JSON.parse(iRbody)
'			    SET oneProduct = strObj.products.get(0)
'				shopifyGoodNo = oneProduct.id
'				shopifyTitle = oneProduct.title
'				shopifyProduct_type = oneProduct.product_type
'				shopifyHandle = oneProduct.handle
'				shopifyUpdated_at = replace(LEFT(oneProduct.updated_at,19),"T"," ")
'				shopifyPublished_at = replace(LEFT(oneProduct.published_at,19),"T"," ")
'				shopifyPublished_scope = oneProduct.published_scope
'				shopifyTags = oneProduct.tags
'				
'				rw shopifyGoodNo&"|"&shopifyTitle&"|"&shopifyProduct_type&"|"&shopifyHandle&"|"&shopifyUpdated_at&"|"&shopifyPublished_scope&"|"&shopifyTags
'				
'				'strStatus	  = strObj.status
'				For i=0 to oneProduct.variants.length-1
'					shopifySKUId = oneProduct.variants.get(i).id
'					shopifySKUtitle = oneProduct.variants.get(i).title
'					shopifySKUprice = oneProduct.variants.get(i).price
'					shopifySKUcompare_at_price = oneProduct.variants.get(i).compare_at_price
'					shopifySKUoption1 = oneProduct.variants.get(i).option1
'					shopifySKUoption2 = oneProduct.variants.get(i).option2
'					shopifySKUoption3 = oneProduct.variants.get(i).option3
'					shopifySKUsku       = oneProduct.variants.get(i).sku
'					shopifySKUbarcode   = oneProduct.variants.get(i).barcode
'					shopifySKUgrams     = oneProduct.variants.get(i).grams
'					shopifySKUinventory_quantity = oneProduct.variants.get(i).inventory_quantity
'					shopifySKUold_inventory_quantity = oneProduct.variants.get(i).old_inventory_quantity
'					shopifySKUposition = oneProduct.variants.get(i).position
'					''shopifySKUinventory_item_id = oneProduct.variants.get(i).inventory_item_id
'					''weight
'					''weight_unit
'					
'					rw shopifySKUId&"|"&shopifySKUtitle&"|"&shopifySKUprice&"|"&shopifySKUcompare_at_price&"|"&shopifySKUoption1&"|"&shopifySKUoption2&"|"&shopifySKUoption3&"|"&shopifySKUsku&"|"&shopifySKUbarcode&"|"&shopifySKUgrams&"|"&shopifySKUgrams
'				Next
'				
'				
''				strSql = ""
''				strSql = strSql & " UPDATE R" & VbCrlf
''				strSql = strSql & " SET lastStatCheckDate = getdate()" & VBCRLF
''				If strStatus = "NEEDS_ACTION" Then					'�ݷ��ε�
''					strSql = strSql & " ,shopifyStatCd = '40'"& VbCRLF
''					strSql = strSql & " ,shopifySellyn = 'N'"& VbCRLF
''				ElseIf strStatus = "APPROVED" Then					'�����ε�
''					strSql = strSql & " ,shopifyStatCd = '7'"& VbCRLF
''					strSql = strSql & " ,shopifySellyn = 'Y'"& VbCRLF
''				Else												'�׿ܴ� ���δ��
''					strSql = strSql & " ,shopifyStatCd = '3'"& VbCRLF
''					strSql = strSql & " ,shopifySellyn = 'N'"& VbCRLF
''				End If
''				strSql = strSql & " ,shopifySkuGoodNo = '" & shopifySKUId & "'" & VbCrlf
''				strSql = strSql & " ,shopifyGoodno = '" & shopifyGoodNo & "'" & VbCrlf
''				strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
''				strSql = strSql & " WHERE R.itemid = " & iitemid
''				strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
''				dbget.execute strSql
''				iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[CHKSTAT]����"
'                SET oneProduct = Noting
'			Set strObj = nothing
'		Else
'			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT]"
'		End If
'	Set objJSON= nothing
'	On Error Goto 0
'End Function

'��ǰ ��ȸ itemArrays
Public Function fnGetShopifyGoodInfo(iitemid, istrGoodNo, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, strStatus, i
	Dim shopifyGoodNo, shopifyTitle, shopifyProduct_type, shopifyHandle, shopifyUpdated_at, shopifyPublished_at, shopifyPublished_scope, shopifyTags
	Dim shopifySKUId, shopifySKUtitle, shopifySKUprice,shopifySKUcompare_at_price,shopifySKUoption1,shopifySKUoption2,shopifySKUoption3,shopifySKUsku,shopifySKUbarcode,shopifySKUgrams,shopifySKUinventory_quantity
	Dim shopifySKUold_inventory_quantity,shopifySKUinventory_item_id,shopifySKUposition
	
	Dim oneProduct
	Dim shopifySellYn, shopifyStatCd, regimagename

	'On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "GET", shopifyAPIURL & "/admin/products.json?ids="&istrGoodNo&"" , False ,shopifySELLERID,shopifyAPIKEY
		objJSON.setRequestHeader "Content-Type", "application/json"
''rw shopifyAPIURL & "/admin/products.json"
		objJSON.Send()
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"utf-8")
			
			Set strObj = JSON.parse(iRbody)
			
			    ''if isEmpty(strObj.products) then
			    if (strObj.products="") then
			        iErrStr = "ERR||"&iitemid&"||�ش� ��ǰ�� �����ϴ�..[ERR-CHKSTAT]"
			    else
			        
    			    SET oneProduct = strObj.products.get(0)
    				shopifyGoodNo = oneProduct.id
    				shopifyTitle = oneProduct.title
    				shopifyProduct_type = oneProduct.product_type
    				shopifyHandle = oneProduct.handle
    				shopifyUpdated_at = replace(LEFT(oneProduct.updated_at,19),"T"," ")
    				shopifyPublished_at = replace(LEFT(oneProduct.published_at,19),"T"," ")
    				shopifyPublished_scope = oneProduct.published_scope
    				shopifyTags = oneProduct.tags
    				
    				''rw shopifyGoodNo&"|"&shopifyTitle&"|"&shopifyProduct_type&"|"&shopifyHandle&"|"&shopifyUpdated_at&"|"&shopifyPublished_scope&"|"&shopifyTags
    				
    				'strStatus	  = strObj.status
    				For i=0 to oneProduct.variants.length-1
    					shopifySKUId = oneProduct.variants.get(i).id
    					shopifySKUtitle = oneProduct.variants.get(i).title
    					shopifySKUprice = oneProduct.variants.get(i).price
    					shopifySKUcompare_at_price = oneProduct.variants.get(i).compare_at_price
    					shopifySKUoption1 = oneProduct.variants.get(i).option1
    					shopifySKUoption2 = oneProduct.variants.get(i).option2
    					shopifySKUoption3 = oneProduct.variants.get(i).option3
    					shopifySKUsku       = oneProduct.variants.get(i).sku
    					shopifySKUbarcode   = oneProduct.variants.get(i).barcode
    					shopifySKUgrams     = oneProduct.variants.get(i).grams
    					shopifySKUinventory_quantity = oneProduct.variants.get(i).inventory_quantity
    					shopifySKUold_inventory_quantity = oneProduct.variants.get(i).old_inventory_quantity
    					shopifySKUposition = oneProduct.variants.get(i).position
    					''shopifySKUinventory_item_id = oneProduct.variants.get(i).inventory_item_id
    					''weight
    					''weight_unit
    					
    					''rw shopifySKUId&"|"&shopifySKUtitle&"|"&shopifySKUprice&"|"&shopifySKUcompare_at_price&"|"&shopifySKUoption1&"|"&shopifySKUoption2&"|"&shopifySKUoption3&"|"&shopifySKUsku&"|"&shopifySKUbarcode&"|"&shopifySKUgrams&"|"&shopifySKUgrams
    				Next
    				
    				if (iitemid="") then
        				''10 01927642 0000
        				if (LEN(shopifySKUsku)=12) and LEFT(shopifySKUsku,2)="10" then
        				    iitemid = Mid(shopifySKUsku,3,11)
        			    elseif (LEN(shopifySKUsku)=14) and LEFT(shopifySKUsku,2)="10" then
        			        iitemid = Mid(shopifySKUsku,3,9)
        				end if
        			end if
    				
    				shopifySellYn = "Y"
    				if (shopifyPublished_at="null") then
        			    shopifyStatCd = "3"
        			else
        			    shopifyStatCd = "7"
        		    end if
    
    				
    				if (iitemid<>"") then
        				strSql = "exec [db_etcmall].[dbo].[usp_TEN_Shopify_UpdateMappingItem] "&iitemid&",'"&shopifyPublished_at&"','"&shopifyUpdated_at&"','"&shopifyGoodNo&"','"&html2db(shopifyHandle)&"',"&shopifySKUprice&","&shopifySKUcompare_at_price&",'"&shopifySellYn&"',"&shopifyStatCd&",'"&html2db(shopifyTitle)&"','"&regimagename&"'"
        
                        rw strSql
        				dbget.execute strSql
        
        				iErrStr =  "OK||"&iitemid&"||[CHKSTAT]����"
    			    end if
                    SET oneProduct = nothing
                END IF
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'��� ��ȸ
Public Function fnshopifySKUGoodNo(iitemid, iitemoption, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/api/v1/products/quantities/forSKUs" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", shopifySELLERID
		objJSON.SetRequestHeader "apiKey", shopifyAPIKEY
		objJSON.SetRequestHeader "locale", shopifyLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			Set strObj = JSON.parse(iRbody)
				For i=0 to strObj.shopifySKUQuantities.length-1
					quantity = strObj.shopifySKUQuantities.get(i).quantity
				Next
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCrlf
				strSql = strSql & " SET quantity = '" & quantity & "'" & VbCrlf
				strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
				strSql = strSql & " WHERE R.itemid = " & iitemid
				strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[CHKQTY]����"
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKQTY]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'��� ����
Public Function fnshopifyEditQuantity(iitemid, iitemoption, imaylimitEa, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/api/v1/products/updateQuantities" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", shopifySELLERID
		objJSON.SetRequestHeader "apiKey", shopifyAPIKEY
		objJSON.SetRequestHeader "locale", shopifyLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess = strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R SET " & VbCrlf
					If imaylimitEa <= 0 Then
						strSql = strSql & "  quantity = quantity - " & imaylimitEa & VbCrlf
					Else
						strSql = strSql & "  quantity = quantity + " & imaylimitEa & VbCrlf
					End If
					strSql = strSql & " ,shopifylastupdate = getdate() "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[EDITQTY]����"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITQTY]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'��� ���� 0����
Public Function fnshopifyEditQuantityZero(iitemid, iitemoption, imaylimitEa, istrParam, byRef iErrStr)
	Dim objJSON, iRbody, strObj, strSql, i, quantity
	On Error Resume Next
	Set objJSON= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objJSON.Open "POST", shopifyAPIURL & "/api/v1/products/updateQuantities" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", shopifySELLERID
		objJSON.SetRequestHeader "apiKey", shopifyAPIKEY
		objJSON.SetRequestHeader "locale", shopifyLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			Set strObj = JSON.parse(iRbody)
				isSuccess = strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R SET " & VbCrlf
					strSql = strSql & " quantity = 0 "
					strSql = strSql & " ,shopifySellyn = 'N' "
					strSql = strSql & " ,accFailCnt = 0 "
					strSql = strSql & " ,shopifylastupdate = getdate() "
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
					strSql = strSql & " WHERE R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||ǰ��ó��"
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITSELLYN]"
		End If
	Set objJSON= nothing
	On Error Goto 0
End Function

'��ǰ ���� ����
Public Function fnshopifyItemPrice(iitemid, iitemoption, istrParam, iOrgprice, irateprice, iMultiplerate, iExchangeRate, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess
	On Error Resume Next
	Set objJSON= CreateObject("Microsoft.XMLHTTP")
	    objJSON.Open "POST", shopifyAPIURL & "/api/v1/products/updatePrice/forProduct" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", shopifySELLERID
		objJSON.SetRequestHeader "apiKey", shopifyAPIKEY
		objJSON.SetRequestHeader "locale", shopifyLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			Set strObj = JSON.parse(iRbody)
				isSuccess	= strObj.STATUS
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET shopifylastupdate = getdate()"
					strSql = strSql & " ,shopifyPrice = '"&irateprice&"' " & VbCrlf
					strSql = strSql & "	,regOrgprice = " & iOrgprice & VbCRLF
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
					strSql = strSql & " ,multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " ,exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[PRICE]����"
				Else
					'iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE] "& db2html(strObj.MESSAGE)
					iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE]����"
				End If
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE]"
		End If
	Set objJSON= nothing
End Function

'��ǰ ���� ����
Public Function fnshopifyItemPriceBySkuNo(iitemid, iitemoption, istrParam, iOrgprice, irateprice, iMultiplerate, iExchangeRate, byRef iErrStr)
	Dim objJSON, jsonDOM, strSql, resultCode, productNo, iRbody
	Dim iMessage, AssignedRow
	Dim strObj, isSuccess
	On Error Resume Next
	Set objJSON= CreateObject("Microsoft.XMLHTTP")
	    objJSON.Open "POST", shopifyAPIURL & "/api/v1/products/updatePrice/forSKUs" , False
		objJSON.setRequestHeader "Content-Type", "application/json"
		objJSON.SetRequestHeader "sellerId", shopifySELLERID
		objJSON.SetRequestHeader "apiKey", shopifyAPIKEY
		objJSON.SetRequestHeader "locale", shopifyLOCALE
		objJSON.Send(istrParam)
		If objJSON.Status = "200" Then
			iRbody = BinaryToText(objJSON.ResponseBody,"euc-kr")
			'response.write iRbody
			'response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess	= strObj.status
				If isSuccess = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE R" & VbCrlf
					strSql = strSql & " SET shopifylastupdate = getdate()"
					strSql = strSql & " ,shopifyPrice = '"&irateprice&"' " & VbCrlf
					strSql = strSql & "	,regOrgprice = " & iOrgprice & VbCRLF
					strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
					strSql = strSql & " ,multiplerate = '"&iMultiplerate&"' " & vbcrlf
					strSql = strSql & " ,exchangeRate = '"&iExchangeRate&"' " & vbcrlf
					strSql = strSql & " FROM db_etcmall.dbo.tbl_shopify_regitem R" & VbCrlf
					strSql = strSql & " where R.itemid = " & iitemid
					strSql = strSql & " and R.itemoption = '"&iitemoption&"' "
					dbget.execute strSql
					iErrStr =  "OK||"&iitemid&"||"&iitemoption&"||[PRICE]����"
				Else
					'iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE] "& db2html(strObj.MESSAGE)
					iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||[PRICE]����"
				End If
			Set strObj = Nothing
		Else
			iErrStr = "ERR||"&iitemid&"||"&iitemoption&"||shopify ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE]"
		End If
	Set objJSON= nothing
End Function

'ī�װ� ���� ���
Public Function fnshopifySubCategory()
	Dim objXML, iRbody, jsResult, strParam, i, j, lp
	Dim attributes, attributeChoices
	Dim colors, sizes, capacities
	Dim depth1Id, depth1Name, depth2Id, depth2Name, depth3Id, depth3Name, isOptional, isMultiSelectable, optFlag, multiFlag
	Dim colorId, colorName
	Dim sizeId, sizeName
	Dim capacitiesId, capacitiesName
	Dim topTemparr

	strSql = ""
	strSql = strSql & " SELECT ROW_NUMBER() OVER (ORDER BY depth3Code ASC) AS RowNum, depth3Code "
	strSql = strSql & " INTO #TBL1 "
	strSql = strSql & " FROM db_etcmall.[dbo].[tbl_shopify_category] "
	strSql = strSql & " GROUP BY depth3Code "
	strSql = strSql & " ORDER BY depth3Code asc "
	dbget.execute strSql

	strSql = ""
	'strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum <= 100 "
	'strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum > 100 and RowNum <= 200 "
	strSql = strSql & " SELECT depth3Code FROM #TBL1 WHERE RowNum > 200 "
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		topTemparr = rsget.getRows
	End If
	rsget.Close

	For lp = 0 To Ubound(topTemparr, 2)
	'	On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		    objXML.Open "GET", shopifyAPIURL & "/api/v1/subCategories/byId/"&topTemparr(0, lp)&"" , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.SetRequestHeader "sellerId", shopifySELLERID
			objXML.SetRequestHeader "apiKey", shopifyAPIKEY
			objXML.SetRequestHeader "locale", shopifyLOCALE
			objXML.Send()
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
			If objXML.Status = "200" Then
				SET jsResult = JSON.parse(iRbody)
					depth1Id	= html2db(jsResult.id)
					depth1Name	= html2db(jsResult.name)

					SET attributes = jsResult.attributes
						For i=0 to attributes.length-1
							depth2Id			= html2db(attributes.get(i).id)
							depth2Name			= html2db(attributes.get(i).name)
							isOptional			= html2db(attributes.get(i).isOptional)
							isMultiSelectable	= html2db(attributes.get(i).isMultiSelectable)
							If isOptional = "True" Then
								optFlag = "Y"
							Else
								optFlag = "N"
							End If

							If isMultiSelectable = "True" Then
								multiFlag = "Y"
							Else
								multiFlag = "N"
							End If

							SET attributeChoices = attributes.get(i).attributeChoices
								For j=0 to attributeChoices.length-1
									depth3Id	= html2db(attributeChoices.get(j).id)
									depth3Name	= html2db(attributeChoices.get(j).name)
									strSql = ""
									strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_shopify_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name, depth3Id, depth3Name, isOptional, isMultiSelectable) VALUES "
									strSql = strSql & " ('attributeChoices', '"&depth1Id&"', '"&depth1Name&"', '"&depth2Id&"', '"&depth2Name&"', '"&depth3Id&"', '"&depth3Name&"', '"&optFlag&"', '"&multiFlag&"') "
									dbget.execute strSql
								Next
	'						rw "----------------------------------------------------------------------"
	'						rw ""
							SET attributeChoices = nothing
						Next
					SET attributes = Nothing

					SET colors = jsResult.colors
						For i=0 to colors.length-1
							colorId		= html2db(colors.get(i).id)
							colorName	= html2db(colors.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_shopify_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('colors', '"&depth1Id&"', '"&depth1Name&"', '"&colorId&"', '"&colorName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET colors = Nothing

					SET sizes = jsResult.sizes
						For i=0 to sizes.length-1
							sizeId		= html2db(sizes.get(i).id)
							sizeName	= html2db(sizes.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_shopify_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('sizes', '"&depth1Id&"', '"&depth1Name&"', '"&sizeId&"', '"&sizeName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET sizes = Nothing

					SET capacities = jsResult.capacities
						For i=0 to capacities.length-1
							capacitiesId	= html2db(capacities.get(i).id)
							capacitiesName	= html2db(capacities.get(i).name)
							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_shopify_subCategory (gubun, depth1Id, depth1Name, depth2Id, depth2Name) VALUES "
							strSql = strSql & " ('capacities', '"&depth1Id&"', '"&depth1Name&"', '"&capacitiesId&"', '"&capacitiesName&"') "
							dbget.execute strSql
	'						rw "----------------------------------------------------------------------"
	'						rw ""
						Next
					SET capacities = Nothing

				SET jsResult = nothing
				response.write "OK||subCategory||����"
			End If
		Set objXML = Nothing
	Next
'	On Error Goto 0
End Function

'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################

'################################################# �� ��� �� �Ķ���� �������� ###############################################
'��� ��ȸ JSON
Public Function fnshopifyQuantitySearchJSON(istrSKUGoodNo)
	Dim strRst
	strRst = ""
	strRst = strRst & "{"
	strRst = strRst & "	""shopifySKUIds"": ["""&istrSKUGoodNo&"""]"
	strRst = strRst & "}"
	fnshopifyQuantitySearchJSON = strRst
End Function

'��� ���� JSON
Public Function fnshopifyQuantityEditJSON(iitemid, iitemoption, iquantity, imaylimitEa, istrSKUGoodNo)
	Dim strRst, strSql
	Dim vLimityn, vLimitNo, vLimitSold, vIsUsing, vSellYn, limitEA, DEFAULTQTY, maySellAvailQty
	Dim oIsusing, oOptsellyn, oOptlimitno, oOptlimitsold

	DEFAULTQTY = 999
	strSql = ""
	strSql = strSql & " SELECT TOP 1 limityn, limitno, limitsold, isusing, sellyn "
	strSql = strSql & " FROM db_item.dbo.tbl_item "
	strSql = strSql & " WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql,dbget,1
	If not rsget.EOF Then
		vLimityn	= rsget("limityn")
		vLimitNo	= rsget("limitno")
		vLimitSold	= rsget("limitsold")
		vIsUsing	= rsget("isusing")
		vSellYn		= rsget("sellyn")
	End If
	rsget.Close

	'iquantity : ���� 10x10 ������ SCM�� ��ϵ� ����
	If vIsUsing <> "Y" OR vSellYn <> "Y" Then
		limitEA = -1 * iquantity
	Else
		If iitemoption = "0000" Then
			If vLimityn = "N" Then
				limitEA = DEFAULTQTY - iquantity
			Else
				maySellAvailQty = vLimitNo - vLimitSold - 5
				If maySellAvailQty < 1 Then
					limitEA = -1 * iquantity
				Else
					limitEA = maySellAvailQty - iquantity
				End If
			End If
		Else
			If vLimityn = "N" Then
				limitEA = DEFAULTQTY - iquantity
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 isusing, optsellyn, optlimitno, optlimitsold "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE itemid = '"&iitemid&"' "
				strSql = strSql & " and itemoption = '"&iitemoption&"' "
				rsget.Open strSql,dbget,1
				If not rsget.EOF Then
					oIsusing		= rsget("isusing")
					oOptsellyn		= rsget("optsellyn")
					oOptlimitno		= rsget("optlimitno")
					oOptlimitsold	= rsget("optlimitsold")
				End If
				rsget.Close
				maySellAvailQty = oOptlimitno - oOptlimitsold - 5
				If (maySellAvailQty < 1) OR (oIsusing <> "Y") OR (oOptsellyn <> "Y") Then
					limitEA = -1 * iquantity
				Else
					limitEA = maySellAvailQty - iquantity
				End If
			End If
		End If
	End If
'	limitEA = 999

	imaylimitEa = limitEA
	strRst = ""
	strRst = strRst & "{"
	strRst = strRst & "	""skuDeltaQuantities"": [{"
	strRst = strRst & "		""shopifySKUId"": """&istrSKUGoodNo&""","
	strRst = strRst & "		""deltaQuantity"": "&limitEA&""
	strRst = strRst & "	}]"
	strRst = strRst & "}"
	fnshopifyQuantityEditJSON = strRst
End Function

%>