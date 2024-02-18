<%
'############################################## ���� �����ϴ� API �Լ� ���� ##############################################
'��ǰ������� ǰ�� �� �׸��ڵ� ��ȸ
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
							itemInfoNtfcAtcCd		= datalist.get(i).itemInfoNtfcAtcCd		'������� ǰ���ڵ� 2 �ڸ�
							itemInfoNtfcAtcNm		= datalist.get(i).itemInfoNtfcAtcNm		'ǰ���
							itemInfoNtfcItmCd		= datalist.get(i).itemInfoNtfcItmCd		'������� �׸��ڵ� 3 �ڸ�
							itemInfoNtfcItmNm		= datalist.get(i).itemInfoNtfcItmNm		'�׸��

							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_marketfor_gosi (itemInfoNtfcAtcCd, itemInfoNtfcAtcNm, itemInfoNtfcItmCd, itemInfoNtfcItmNm) VALUES "
							sqlStr = sqlStr & " ('"& itemInfoNtfcAtcCd &"', '"& itemInfoNtfcAtcNm &"', '"& itemInfoNtfcItmCd &"', '"& itemInfoNtfcItmNm &"') "
							dbget.execute sqlStr
						Next
					Set datalist = nothing
					rw "GOSI INSERT END"
				Else
					rw "�����ڵ� 0000�� �ƴ� : " & returnCode
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ������� ǰ�� �� �׸��ڵ� ��ȸ
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
							mfItemClasfCtgId		= mfItemClasfCtgList.get(i).mfItemClasfCtgId		'��������ǰ�з�ī�װ� ID
							rhrnMfItemClasfCtgId	= mfItemClasfCtgList.get(i).rhrnMfItemClasfCtgId	'������������ǰ�з�ī�װ� ID
							itemClasfCtgLvlVal		= mfItemClasfCtgList.get(i).itemClasfCtgLvlVal		'��ǰ�з�ī�װ�������
							mfItemClasfCtgNm		= mfItemClasfCtgList.get(i).mfItemClasfCtgNm		'��������ǰ�з�ī�װ���
							mfItemLclsCtgId			= mfItemClasfCtgList.get(i).mfItemLclsCtgId			'��ǰ��з�ī�װ� ID
							mfItemLclsCtgNm			= mfItemClasfCtgList.get(i).mfItemLclsCtgNm			'��ǰ��з�ī�װ���
							mfItemMclsCtgId			= mfItemClasfCtgList.get(i).mfItemMclsCtgId			'��ǰ�ߺз�ī�װ� ID
							mfItemMclsCtgNm			= mfItemClasfCtgList.get(i).mfItemMclsCtgNm			'��ǰ�ߺз�ī�װ���
							mfItemSclsCtgId			= mfItemClasfCtgList.get(i).mfItemSclsCtgId			'��ǰ�Һз�ī�װ� ID
							mfItemSclsCtgNm			= mfItemClasfCtgList.get(i).mfItemSclsCtgNm			'��ǰ�Һз�ī�װ���
							mfItemDclsCtgId			= mfItemClasfCtgList.get(i).mfItemDclsCtgId			'��ǰ���з�ī�װ� ID
							mfItemDclsCtgNm			= mfItemClasfCtgList.get(i).mfItemDclsCtgNm			'��ǰ���з�ī�װ���
							mfItemVclsCtgId			= mfItemClasfCtgList.get(i).mfItemVclsCtgId			'��ǰ�󼼺з�ī�װ� ID
							mfItemVclsCtgNm			= mfItemClasfCtgList.get(i).mfItemVclsCtgNm			'��ǰ�󼼺з�ī�װ���
							mfItemTclsCtgId			= mfItemClasfCtgList.get(i).mfItemTclsCtgId			'��ǰ�󼼼��з�ī�װ� ID
							mfItemTclsCtgNm			= mfItemClasfCtgList.get(i).mfItemTclsCtgNm			'��ǰ�󼼼��з�ī�װ���
							mfDspCtgId				= mfItemClasfCtgList.get(i).mfDspCtgId
							useYn					= mfItemClasfCtgList.get(i).useYn
							expsrSeq				= mfItemClasfCtgList.get(i).expsrSeq
							Set itemClassfNtfcPosbList 	= mfItemClasfCtgList.get(i).itemClassfNtfcPosbList
								For j=0 to itemClassfNtfcPosbList.length-1
									itemInfoNtfcAtcCd = ""
									itemInfoNtfcAtcNm = ""
									itemInfoNtfcAtcCd = itemClassfNtfcPosbList.get(j).itemInfoNtfcAtcCd	'��ǰ�������ǰ���ڵ�
									itemInfoNtfcAtcNm = itemClassfNtfcPosbList.get(j).itemInfoNtfcAtcNm	'ǰ���

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
					rw "�����ڵ� 0000�� �ƴ� : " & returnCode
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

%>