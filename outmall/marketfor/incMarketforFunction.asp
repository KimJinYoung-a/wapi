<%
'############################################## 실제 수행하는 API 함수 모음 ##############################################
'상품정보고시 품목 및 항목코드 조회
Public Function fnMarketforGetGosiCode(iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, datalist, i, sqlStr
	Dim itemInfoNtfcAtcCd, itemInfoNtfcAtcNm, itemInfoNtfcItmCd, itemInfoNtfcItmNm

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIURL & "/ntfc", false
		objXML.setRequestHeader "Authorization", "Basic " & APIkey
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			If iresponseJson = "Y" Then
				response.write iRbody
				response.end
			End If

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "0000" Then
					sqlStr = " DELETE FROM db_etcmall.dbo.tbl_marketfor_gosi "
					dbget.execute sqlStr
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							itemInfoNtfcAtcCd		= datalist.get(i).itemInfoNtfcAtcCd		'정보고시 품목코드 2 자리
							itemInfoNtfcAtcNm		= datalist.get(i).itemInfoNtfcAtcNm		'품목명
							itemInfoNtfcItmCd		= datalist.get(i).itemInfoNtfcItmCd		'정보고시 항목코드 3 자리
							itemInfoNtfcItmNm		= datalist.get(i).itemInfoNtfcItmNm		'항목명

							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_marketfor_gosi (itemInfoNtfcAtcCd, itemInfoNtfcAtcNm, itemInfoNtfcItmCd, itemInfoNtfcItmNm) VALUES "
							sqlStr = sqlStr & " ('"& itemInfoNtfcAtcCd &"', '"& itemInfoNtfcAtcNm &"', '"& itemInfoNtfcItmCd &"', '"& itemInfoNtfcItmNm &"') "
							dbget.execute sqlStr
						Next
					Set datalist = nothing
					rw "GOSI INSERT END"
				Else
					rw "리턴코드 0000이 아님 : " & returnCode
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품정보고시 품목 및 항목코드 조회
Public Function fnMarketforGetClsCateCode(iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, mfItemClasfCtgList, i, j, sqlStr
	Dim mfItemClasfCtgId, rhrnMfItemClasfCtgId, itemClasfCtgLvlVal, mfItemClasfCtgNm, mfItemLclsCtgId, mfItemLclsCtgNm, mfItemMclsCtgId, mfItemMclsCtgNm, mfItemSclsCtgId, mfItemSclsCtgNm, mfItemDclsCtgId, mfItemDclsCtgNm, mfItemVclsCtgId, mfItemVclsCtgNm, mfItemTclsCtgId, mfItemTclsCtgNm, mfDspCtgId, useYn, expsrSeq
	Dim itemClassfNtfcPosbList, itemInfoNtfcAtcCd, itemInfoNtfcAtcNm

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIURL & "/ctg", false
		objXML.setRequestHeader "Authorization", "Basic " & APIkey
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			'response.write iRbody
			If iresponseJson = "Y" Then
				response.write iRbody
				response.end
			End If

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.code
				If returnCode = "0000" Then
					sqlStr = " DELETE FROM db_etcmall.[dbo].[tbl_marketfor_clscategory] "
					dbget.execute sqlStr

					sqlStr = " DELETE FROM db_etcmall.[dbo].[tbl_marketfor_metainfo] "
					dbget.execute sqlStr
					Set mfItemClasfCtgList = strObj.data.mfItemClasfCtgList
						For i=0 to mfItemClasfCtgList.length-1
							mfItemClasfCtgId		= mfItemClasfCtgList.get(i).mfItemClasfCtgId		'마켓포상품분류카테고리 ID
							rhrnMfItemClasfCtgId	= mfItemClasfCtgList.get(i).rhrnMfItemClasfCtgId	'상위마켓포상품분류카테고리 ID
							itemClasfCtgLvlVal		= mfItemClasfCtgList.get(i).itemClasfCtgLvlVal		'상품분류카테고리레벨값
							mfItemClasfCtgNm		= mfItemClasfCtgList.get(i).mfItemClasfCtgNm		'마켓포상품분류카테고리명
							mfItemLclsCtgId			= mfItemClasfCtgList.get(i).mfItemLclsCtgId			'상품대분류카테고리 ID
							mfItemLclsCtgNm			= mfItemClasfCtgList.get(i).mfItemLclsCtgNm			'상품대분류카테고리명
							mfItemMclsCtgId			= mfItemClasfCtgList.get(i).mfItemMclsCtgId			'상품중분류카테고리 ID
							mfItemMclsCtgNm			= mfItemClasfCtgList.get(i).mfItemMclsCtgNm			'상품중분류카테고리명
							mfItemSclsCtgId			= mfItemClasfCtgList.get(i).mfItemSclsCtgId			'상품소분류카테고리 ID
							mfItemSclsCtgNm			= mfItemClasfCtgList.get(i).mfItemSclsCtgNm			'상품소분류카테고리명
							mfItemDclsCtgId			= mfItemClasfCtgList.get(i).mfItemDclsCtgId			'상품세분류카테고리 ID
							mfItemDclsCtgNm			= mfItemClasfCtgList.get(i).mfItemDclsCtgNm			'상품세분류카테고리명
							mfItemVclsCtgId			= mfItemClasfCtgList.get(i).mfItemVclsCtgId			'상품상세분류카테고리 ID
							mfItemVclsCtgNm			= mfItemClasfCtgList.get(i).mfItemVclsCtgNm			'상품상세분류카테고리명
							mfItemTclsCtgId			= mfItemClasfCtgList.get(i).mfItemTclsCtgId			'상품상세세분류카테고리 ID
							mfItemTclsCtgNm			= mfItemClasfCtgList.get(i).mfItemTclsCtgNm			'상품상세세분류카테고리명
							mfDspCtgId				= mfItemClasfCtgList.get(i).mfDspCtgId
							useYn					= mfItemClasfCtgList.get(i).useYn
							expsrSeq				= mfItemClasfCtgList.get(i).expsrSeq
							Set itemClassfNtfcPosbList 	= mfItemClasfCtgList.get(i).itemClassfNtfcPosbList
								For j=0 to itemClassfNtfcPosbList.length-1
									itemInfoNtfcAtcCd = ""
									itemInfoNtfcAtcNm = ""
									itemInfoNtfcAtcCd = itemClassfNtfcPosbList.get(j).itemInfoNtfcAtcCd	'상품정보고시품목코드
									itemInfoNtfcAtcNm = itemClassfNtfcPosbList.get(j).itemInfoNtfcAtcNm	'품목명

									sqlStr = ""
									sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_marketfor_metainfo] (mfItemClasfCtgId, itemInfoNtfcAtcCd, itemInfoNtfcAtcNm) VALUES "
									sqlStr = sqlStr & " ('"& mfItemClasfCtgId &"', '"& itemInfoNtfcAtcCd &"', '"& itemInfoNtfcAtcNm &"') "
									dbget.execute sqlStr
								Next
							Set itemClassfNtfcPosbList = nothing
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_marketfor_clscategory] (mfItemClasfCtgId, rhrnMfItemClasfCtgId, itemClasfCtgLvlVal, mfItemClasfCtgNm, mfItemLclsCtgId, mfItemLclsCtgNm, mfItemMclsCtgId, mfItemMclsCtgNm, mfItemSclsCtgId, mfItemSclsCtgNm, mfItemDclsCtgId, mfItemDclsCtgNm, mfItemVclsCtgId, mfItemVclsCtgNm, mfItemTclsCtgId, mfItemTclsCtgNm, mfDspCtgId, useYn, expsrSeq) VALUES "
							sqlStr = sqlStr & " ('"& mfItemClasfCtgId &"', '"& rhrnMfItemClasfCtgId &"', '"& itemClasfCtgLvlVal &"', '"& mfItemClasfCtgNm &"', '"& mfItemLclsCtgId &"', '"& mfItemLclsCtgNm &"', '"& mfItemMclsCtgId &"', '"& mfItemMclsCtgNm &"', '"& mfItemSclsCtgId &"', '"& mfItemSclsCtgNm &"', '"& mfItemDclsCtgId &"', '"& mfItemDclsCtgNm &"', '"& mfItemVclsCtgId &"', '"& mfItemVclsCtgNm &"', '"& mfItemTclsCtgId &"', '"& mfItemTclsCtgNm &"', '"& mfDspCtgId &"', '"& useYn &"', '"& expsrSeq &"') "
							dbget.execute sqlStr
						Next
					Set mfItemClasfCtgList = nothing
					rw "ClsCateGory INSERT END"
				Else
					rw "리턴코드 0000이 아님 : " & returnCode
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

%>