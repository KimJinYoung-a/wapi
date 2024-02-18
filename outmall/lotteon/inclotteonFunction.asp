<%
'############################################## ���� �����ϴ� API �Լ� ���� ���� ############################################
'�Ǹ��� �����/��ǰ�� ����Ʈ��ȸ
Public Function fnlotteonDVPView
    Dim objXML, iRbody, strObj, returnCode, datalist, i
	Dim dvpNo, dvpTypCd, dvpNm, zipNo, zipAddr, dtlAddr, stnmZipNo, stnmZipAddr, stnmDtlAddr, rpbtrNm, mphnNatnNoCd, mphnNo, telNatnNoCd, telNo, lrtrNo, useYn
	Dim obj

	Set obj = jsObject()
		obj("afflTrCd") = afflTrCd
		strParam = obj.jsString
	Set obj = nothing

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/contract/v1/dvp/getDvpListSr", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[DVP] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.returnCode
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							dvpNo			= datalist.get(i).dvpNo			''PLE&PLE'+�������ȣ
							dvpTypCd		= datalist.get(i).dvpTypCd		'����������ڵ� [ȸ����:01, �����:02]
							dvpNm			= datalist.get(i).dvpNm			'�������
							zipNo			= datalist.get(i).zipNo			'�����ȣ
							zipAddr			= datalist.get(i).zipAddr		'�����ּ�
							dtlAddr			= datalist.get(i).dtlAddr		'���ּ�
							stnmZipNo		= datalist.get(i).stnmZipNo		'���θ�����ȣ
							stnmZipAddr		= datalist.get(i).stnmZipAddr	'���θ�����ּ�
							stnmDtlAddr		= datalist.get(i).stnmDtlAddr	'�������ּ�
							rpbtrNm			= datalist.get(i).rpbtrNm		'����ڸ�
							mphnNatnNoCd	= datalist.get(i).mphnNatnNoCd	'�޴��������ڵ� [default:'82']
							mphnNo			= datalist.get(i).mphnNo		'�޴�����ȣ
							telNatnNoCd		= datalist.get(i).telNatnNoCd	'����ó�����ڵ� [defaul:'82']
							telNo			= datalist.get(i).telNo			'����ó
							lrtrNo			= datalist.get(i).lrtrNo		'??�����ŷ�ó��ȣ
							useYn			= datalist.get(i).useYn			'��뿩�� [default:'Y']

							If useYn = "Y" Then
								rw "'PLE&PLE'+�������ȣ : " & dvpNo
								rw "����������ڵ� [ȸ����:01, �����:02] : " & dvpTypCd
								rw "������� : " & dvpNm
								rw "�����ȣ : " & zipNo
								rw "�����ּ� : " & zipAddr
								rw "���ּ� : " & dtlAddr
								rw "���θ�����ȣ : " & stnmZipNo
								rw "���θ�����ּ� : " & stnmZipAddr
								rw "�������ּ� : " & stnmDtlAddr
								rw "����ڸ� : " & rpbtrNm
								rw "�޴��������ڵ� : " & mphnNatnNoCd
								rw "�޴�����ȣ : " & mphnNo
								rw "����ó�����ڵ� : " & telNatnNoCd
								rw "����ó : " & telNo
								rw "??�����ŷ�ó��ȣ : " & lrtrNo
								rw "��뿩�� : " & useYn
								rw "--------------------------------------"
							End If
						Next
					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||����[�׷���ȸ] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�Ӽ� �⺻ ��ȸ
Public Function fnlotteonAttrView(skip, ihasNext)
    Dim objXML, iRbody, strObj, returnCode, itemList, attr_val_list, i, j, strSql
	Dim attr_id, attr_val_frm_cd, use_yn, mod_date, attr_nm, attr_pi_type, attr_disp_nm
	Dim attr_ref_val1, attr_ref_val2, attr_val_nm, attr_val_disp_nm, use_yn2, attr_val_id
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIATTRURL & "/cheetah/econCheetah.ecn?job=cheetahAttr&skip="&skip&"&limit=500&sort=sort_2&direction=asc", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�׷���ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList
				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						attr_id			= ""
						attr_id			= itemList.get(i).data.attr_id					'�Ӽ����� ID
						attr_val_frm_cd = itemList.get(i).data.attr_val_frm_cd			'�Ӽ������ı����ڵ�
						use_yn			= itemList.get(i).data.use_yn					'��뿩��
						mod_date		= itemList.get(i).data.mod_date					'�����Ͻ�
						attr_nm			= itemList.get(i).data.attr_nm					'�Ӽ�������
						attr_pi_type	= itemList.get(i).data.attr_pi_type				'�Ӽ���ǰ���±���
						attr_disp_nm	= itemList.get(i).data.attr_disp_nm				'�Ӽ��������ø�

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_Attribute_Ins] '"&attr_id&"', '"&attr_val_frm_cd&"', '"&use_yn&"', '"&mod_date&"' " & VBCRLF
						strSql = strSql & " ,'"&attr_nm&"' ,'"&attr_pi_type&"' ,'"&attr_disp_nm&"'  "
   						dbget.Execute(strSql)

						' rw "�Ӽ����� ID : " & attr_id
						' rw "�Ӽ������ı����ڵ� : " & attr_val_frm_cd
						' rw "��뿩�� : " & use_yn
						' rw "�����Ͻ� : " & mod_date
						' rw "�Ӽ������� : " & attr_nm
						' rw "�Ӽ���ǰ���±��� : " & attr_pi_type
						' rw "�Ӽ��������ø� : " & attr_disp_nm
						' rw "--------------------------------------"
						Set attr_val_list = itemList.get(i).data.attr_val_list
						 	For j = 0 to attr_val_list.length-1
								attr_ref_val1		= ""
								attr_ref_val2		= ""
								attr_val_nm			= ""
								attr_val_disp_nm	= ""
								use_yn2				= ""
								attr_val_id			= ""

								attr_ref_val1		= attr_val_list.get(j).attr_ref_val1		'�Ӽ�������1 (������ ��� ����迭, �������� ��� �����迭)
								attr_ref_val2		= attr_val_list.get(j).attr_ref_val2		'�Ӽ�������2 (�������� ��� ���� ������ ���̴� ����)
								attr_val_nm			= attr_val_list.get(j).attr_val_nm			'�Ӽ�����
								attr_val_disp_nm	= attr_val_list.get(j).attr_val_disp_nm		'�Ӽ������ø�
								use_yn2				= attr_val_list.get(j).use_yn				'��뿩��
								attr_val_id			= attr_val_list.get(j).attr_val_id			'�Ӽ���ID

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_AttributeValue_Ins] '"&attr_id&"', '"&attr_val_id&"', '"&attr_val_nm&"' " & VBCRLF
								strSql = strSql & " ,'"&attr_val_disp_nm&"' ,'"&attr_ref_val1&"' ,'"&attr_ref_val2&"','"&use_yn2&"'  "
								dbget.Execute(strSql)
								' rw "�Ӽ�������1 : " & attr_ref_val1
								' rw "�Ӽ�������2 : " & attr_ref_val2
								' rw "�Ӽ����� : " & attr_val_nm
								' rw "�Ӽ������ø� : " & attr_val_disp_nm
								' rw "��뿩�� : " & use_yn2
								' rw "�Ӽ���ID : " & attr_val_id
						 	Next
						Set attr_val_list = nothing
						' rw "###################################################"
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=ATTRVIEW&rSkip=0
						' rw "###################################################"
					Next
					rw "�Ӽ� " & skip & " ���� " & itemList.length & " �� ���"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					rw "���̻� ����"
					ihasNext = "N"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'����ī�װ� ��ȸ
Public Function fnlotteonDispCateView(skip, ihasNext)
    Dim objXML, iRbody, strObj, returnCode, itemList, datalist, i, j, k, strSql
	Dim disp_cat_id, disp_cat_nm, disp_cat_desc, upr_disp_cat_id, depth_no, leaf_yn, prio_rnk, mall_dvs_cd, disp_yn, mod_date, use_yn

	If skip = "skip" Then
		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_lotteon_DispCategory] "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_DispCategory_Ins] "
		dbget.Execute(strSql)

		rw "����ī�װ� ���� �Ϸ�"
		Exit Function
	End If

'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIATTRURL & "/cheetah/econCheetah.ecn?job=cheetahDisplayCategory&skip="&skip&"&limit=500&sort=sort_2&direction=asc", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�׷���ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList

				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						disp_cat_id		= itemList.get(i).data.disp_cat_id			'����ī�װ�ID
						disp_cat_nm 	= itemList.get(i).data.disp_cat_nm			'����ī�װ���
						disp_cat_desc	= itemList.get(i).data.disp_cat_desc		'����ī�װ�����
						upr_disp_cat_id	= itemList.get(i).data.upr_disp_cat_id		'��������ī�װ���ȣ
						depth_no		= itemList.get(i).data.depth_no				'���̹�ȣ
						leaf_yn			= itemList.get(i).data.leaf_yn				'��������
						prio_rnk		= itemList.get(i).data.prio_rnk				'�켱���� (����ī�װ� �켱����)
						mall_dvs_cd		= itemList.get(i).data.mall_dvs_cd			'�������ڵ�
						disp_yn			= itemList.get(i).data.disp_yn				'���ÿ���
						mod_date		= itemList.get(i).data.mod_date				'�����Ͻ�
						use_yn			= itemList.get(i).data.use_yn				'��뿩��

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_TmpDispCategory_Ins] '"&disp_cat_id&"', '"&disp_cat_nm&"', '"&disp_cat_desc&"', '"&upr_disp_cat_id&"' " & VBCRLF
						strSql = strSql & " ,'"&depth_no&"' ,'"&leaf_yn&"' ,'"&prio_rnk&"' ,'"&mall_dvs_cd&"' ,'"&disp_yn&"' ,'"&mod_date&"' ,'"&use_yn&"'  "
   						dbget.Execute(strSql)

						'����ī�װ� ����� ����
						'1. http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=DISPCATE&rSkip=0
						'2. �Ʒ� Url ȣ���Ͽ� ī�װ� ����
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=DISPCATE&rSkip=skip

						' rw "����ī�װ�ID : " & disp_cat_id
						' rw "����ī�װ��� : " & disp_cat_nm
						' rw "����ī�װ����� : " & disp_cat_desc
						' rw "��������ī�װ���ȣ : " & upr_disp_cat_id
						' rw "���̹�ȣ : " & depth_no
						' rw "�������� : " & leaf_yn
						' rw "�켱���� : " & prio_rnk
						' rw "�������ڵ� : " & mall_dvs_cd
						' rw "���ÿ��� : " & disp_yn
						' rw "�����Ͻ� : " & mod_date
						' rw "��뿩�� : " & use_yn
						' rw "---------------------------------------------------"
					Next
					rw "����ī�װ� " & skip & " ���� " & itemList.length & " �� ���"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					rw "���̻� ����"
					ihasNext = "N"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'ǥ��ī�װ� ��ȸ
Public Function fnlotteonStdCateView(skip, ihasNext)
    Dim objXML, iRbody, strObj, returnCode, itemList, datalist, i, j, k, strSql
	Dim std_cat_id, std_cat_nm, std_cat_desc, upr_std_cat_id, depth_no, leaf_yn, prio_rnk, tdf_cd, age_limit_cd, deduct_cult_yn, aband_pickup_yn, exch_money_yn, counsel_prod_yn
	Dim mod_date, use_yn, chl_athn, chl_cfm, chl_sups, cmcn_athn, cmcn_reg, cmcn_tntt, elc_athn, elc_cfm, elc_sups, life_athn, life_cfm, life_sups, life_std, chem_life, chem_bioc, etc

	Dim disp_list, dstd_cat_id, dmall_dvs_cd, ddisp_cat_id
	Dim attr_list, astd_cat_id, aattr_id, aprio_rnk
	Dim pd_itms_list

	If skip = "skip" Then
		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_lotteon_StdCategory] "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_StdCategory_Ins] "
		dbget.Execute(strSql)

		rw "ǥ��ī�װ� ���� �Ϸ�"
		response.end
		Exit Function
	End If

'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIATTRURL & "/cheetah/econCheetah.ecn?job=cheetahStandardCategory&skip="&skip&"&limit=500&sort=sort_2&direction=asc", false		'sort_1 = ������ / sort_2 = ����ī�װ�ID
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[ǥ����ȸ] " & html2db(Err.Description)
			Exit Function
		End If
		' rw objXML.Status
		' rw BinaryToText(objXML.ResponseBody,"utf-8")
		' response.end

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList

				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						std_cat_id		= itemList.get(i).data.std_cat_id			'ǥ��ī�װ�ID
'						rw "ǥ��ī�װ�ID : " & std_cat_id

						Set disp_list = itemList.get(i).data.disp_list				'����ī�װ� ����Ʈ
							For j = 0 to disp_list.length-1
								dstd_cat_id 	= ""
								dmall_dvs_cd 	= ""
								ddisp_cat_id 	= ""

								dstd_cat_id		= disp_list.get(j).std_cat_id			'ǥ��ī�װ�ID
								dmall_dvs_cd	= disp_list.get(j).mall_dvs_cd			'�������ڵ�
								ddisp_cat_id	= disp_list.get(j).disp_cat_id			'����ī�װ�ID

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_StdCategoryDisp_Ins] '"&dstd_cat_id&"', '"&ddisp_cat_id&"', '"&dmall_dvs_cd&"' "
								dbget.Execute(strSql)
'								rw " l ǥ��ī�װ�ID : " & dstd_cat_id
'								rw " l �������ڵ� : " & dmall_dvs_cd
'								rw " l ����ī�װ�ID : " & ddisp_cat_id
'								rw "lllllllllllllllllllllllllllllllllllllllllllllllllll"
							Next
						Set disp_list = nothing

						Set attr_list = itemList.get(i).data.attr_list				'�Ӽ����� ����Ʈ
							For k = 0 to attr_list.length-1
								astd_cat_id	= ""
								aattr_id	= ""
								aprio_rnk	= ""

								astd_cat_id	= attr_list.get(k).std_cat_id				'ǥ��ī�װ�ID
								aattr_id	= attr_list.get(k).attr_id					'�Ӽ�����ID
								aprio_rnk	= attr_list.get(k).prio_rnk					'�켱���� (ǥ��ī�װ� - �Ӽ����� ���ο� ���� �켱����)

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_StdCategoryAttr_Ins] '"&astd_cat_id&"', '"&aattr_id&"', '"&aprio_rnk&"' "
								dbget.Execute(strSql)
'								rw " a ǥ��ī�װ�ID : " & astd_cat_id
'								rw " a �Ӽ�����ID : " & aattr_id
'								rw " a �켱���� : " & aprio_rnk
'								rw "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
							Next
						Set attr_list = nothing
'						Set pd_itms_list = itemList.get(i).data.pd_itms_list		'��ǰǰ���� ���� ����Ʈ
						std_cat_nm		= itemList.get(i).data.std_cat_nm			'ǥ��ī�װ���
						std_cat_desc	= itemList.get(i).data.std_cat_desc			'ǥ��ī�װ�����
						upr_std_cat_id	= itemList.get(i).data.upr_std_cat_id		'����ǥ��ī�װ���ȣ
						depth_no		= itemList.get(i).data.depth_no				'���̹�ȣ
						leaf_yn			= itemList.get(i).data.leaf_yn				'��������
						prio_rnk		= itemList.get(i).data.prio_rnk				'�켱���� (ǥ��ī�װ� �켱����)
						tdf_cd			= itemList.get(i).data.tdf_cd				'���������ڵ�
						age_limit_cd	= itemList.get(i).data.age_limit_cd			'���������ڵ�
						deduct_cult_yn	= itemList.get(i).data.deduct_cult_yn		'������ȭ�� ��������
						aband_pickup_yn	= itemList.get(i).data.aband_pickup_yn		'���� ���ſ���
						exch_money_yn	= itemList.get(i).data.exch_money_yn		'ȯ�ݼ� ����
						counsel_prod_yn	= itemList.get(i).data.counsel_prod_yn		'������� ��ǰ����
						mod_date		= itemList.get(i).data.mod_date				'�����Ͻ�
						use_yn			= itemList.get(i).data.use_yn				'��뿩��
						chl_athn		= itemList.get(i).data.chl_athn				'[�����ǰ]��������
						chl_cfm			= itemList.get(i).data.chl_cfm				'[�����ǰ]����Ȯ��
						chl_sups		= itemList.get(i).data.chl_sups				'[�����ǰ]���������ռ�Ȯ��
						cmcn_athn		= itemList.get(i).data.cmcn_athn			'[�����ű�����]��������
						cmcn_reg		= itemList.get(i).data.cmcn_reg				'[�����ű�����]���յ��
						cmcn_tntt		= itemList.get(i).data.cmcn_tntt			'[�����ű�����]��������
						elc_athn		= itemList.get(i).data.elc_athn				'[�����ǰ]��������
						elc_cfm			= itemList.get(i).data.elc_cfm				'[�����ǰ]����Ȯ��
						elc_sups		= itemList.get(i).data.elc_sups				'[�����ǰ]���������ռ�Ȯ��
						life_athn		= itemList.get(i).data.life_athn			'[��Ȱ��ǰ]��������
						life_cfm		= itemList.get(i).data.life_cfm				'[��Ȱ��ǰ]����Ȯ��
						life_sups		= itemList.get(i).data.life_sups			'[��Ȱ��ǰ]���������ռ�Ȯ��
						life_std		= itemList.get(i).data.life_std				'[��Ȱ��ǰ]���������ؼ�
						chem_life		= itemList.get(i).data.chem_life			'[ȭ����ǰ] ��Ȱȭ����ǰ ������������Ȯ�νŰ��ȣ / ���ι�ȣ
						chem_bioc		= itemList.get(i).data.chem_bioc			'[ȭ����ǰ] �������ǰ ���ι�ȣ
						etc				= itemList.get(i).data.etc					'��Ÿ

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_TmpStdCategory_Ins] '"&std_cat_id&"', '"&std_cat_nm&"', '"&std_cat_desc&"', '"&upr_std_cat_id&"' " & VBCRLF
						strSql = strSql & " ,'"&depth_no&"' ,'"&leaf_yn&"' ,'"&prio_rnk&"' ,'"&tdf_cd&"' ,'"&age_limit_cd&"' ,'"&deduct_cult_yn&"' ,'"&aband_pickup_yn&"' ,'"&exch_money_yn&"' " & VBCRLF
						strSql = strSql & " ,'"&counsel_prod_yn&"' ,'"&mod_date&"' ,'"&use_yn&"' ,'"&chl_athn&"' ,'"&chl_cfm&"' ,'"&chl_sups&"' ,'"&cmcn_athn&"' ,'"&cmcn_reg&"' " & VBCRLF
						strSql = strSql & " ,'"&cmcn_tntt&"' ,'"&elc_athn&"' ,'"&elc_cfm&"' ,'"&elc_sups&"' ,'"&life_athn&"' ,'"&life_cfm&"' ,'"&life_sups&"' ,'"&life_std&"' " & VBCRLF
						strSql = strSql & " ,'"&chem_life&"' ,'"&chem_bioc&"' ,'"&etc&"' "
   						dbget.Execute(strSql)

						'ǥ��ī�װ� ����� ����
						'1.	http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=STDCATE&rSkip=0

						'2. �Ʒ� Url ȣ���Ͽ� ī�װ� ����
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=STDCATE&rSkip=skip

						' rw "ǥ��ī�װ��� : " & std_cat_nm
						' rw "ǥ��ī�װ����� : " & std_cat_desc
						' rw "����ǥ��ī�װ���ȣ : " & upr_std_cat_id
						' rw "���̹�ȣ : " & depth_no
						' rw "�������� : " & leaf_yn
						' rw "�켱���� : " & prio_rnk
						' rw "���������ڵ� : " & tdf_cd
						' rw "���������ڵ� : " & age_limit_cd
						' rw "������ȭ�� �������� : " & deduct_cult_yn
						' rw "���� ���ſ��� : " & aband_pickup_yn
						' rw "ȯ�ݼ� ���� : " & exch_money_yn
						' rw "������� ��ǰ���� : " & counsel_prod_yn
						' rw "�����Ͻ� : " & mod_date
						' rw "��뿩�� : " & use_yn
						' rw "[�����ǰ]�������� : " & chl_athn
						' rw "[�����ǰ]����Ȯ�� : " & chl_cfm
						' rw "[�����ǰ]���������ռ�Ȯ�� : " & chl_sups
						' rw "[�����ű�����]�������� : " & cmcn_athn
						' rw "[�����ű�����]���յ�� : " & cmcn_reg
						' rw "[�����ű�����]�������� : " & cmcn_tntt
						' rw "[�����ǰ]�������� : " & elc_athn
						' rw "[�����ǰ]����Ȯ�� : " & elc_cfm
						' rw "[�����ǰ]���������ռ�Ȯ�� : " & elc_sups
						' rw "[��Ȱ��ǰ]�������� : " & life_athn
						' rw "[��Ȱ��ǰ]����Ȯ�� : " & life_cfm
						' rw "[��Ȱ��ǰ]���������ռ�Ȯ�� : " & life_sups
						' rw "[��Ȱ��ǰ]���������ؼ� : " & life_std
						' rw "[ȭ����ǰ] ��Ȱȭ����ǰ ������������Ȯ�νŰ��ȣ / ���ι�ȣ : " & chem_life
						' rw "[ȭ����ǰ] �������ǰ ���ι�ȣ : " & chem_bioc
						' rw "��Ÿ : " & etc
						' rw "###################################################"
						' rw "###################################################"
					Next
					rw "ǥ��ī�װ� " & skip & " ���� " & itemList.length & " �� ���"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					ihasNext = "N"
					rw "���̻� ����"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�귣�� ��ȸ
Public Function fnlotteonBrandView(skip, limit)
    Dim objXML, iRbody, strObj, returnCode, itemList, datalist, i, j, k, strSql
	Dim brnd_sct_cd, type_keyword, use_yn, mod_date, brnd_nm_kr, brnd_logo_img_url1, brnd_nm_etc1, brnd_nm_etc2, brnd_nm, brnd_nm_en, brnd_id, brnd_nm_main_cd, brnd_desc, synonym_keyword

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIATTRURL & "/cheetah/econCheetah.ecn?job=cheetahBrnd&skip="&skip&"&limit="&limit&"&sort=2&direction=asc", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�귣����ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList

				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						brnd_sct_cd			= html2db(itemList.get(i).data.brnd_sct_cd)			'�귣���������ڵ� (��ǰ�����: ��ǰ�귣��)
						type_keyword		= html2db(itemList.get(i).data.typo_keyword)		'�귣���Ÿ����
						use_yn				= itemList.get(i).data.use_yn						'��뿩��
						mod_date			= itemList.get(i).data.mod_date						'�����Ͻ�
						brnd_nm_kr			= html2db(itemList.get(i).data.brnd_nm_kr)			'�귣���ѱ۸�
						brnd_logo_img_url1	= itemList.get(i).data.brnd_logo_img_url1			'�귣��ΰ�ù��°�̹���URL
						brnd_nm_etc1		= html2db(itemList.get(i).data.brnd_nm_etc1)		'�귣���Ÿ��1
						brnd_nm_etc2		= html2db(itemList.get(i).data.brnd_nm_etc2)		'�귣���Ÿ��2
						brnd_nm				= html2db(itemList.get(i).data.brnd_nm)				'�귣���ǥ��
						brnd_nm_en			= html2db(itemList.get(i).data.brnd_nm_en)			'�귣�念����
						brnd_id				= html2db(itemList.get(i).data.brnd_id)				'�귣��ID
						brnd_nm_main_cd		= html2db(itemList.get(i).data.brnd_nm_main_cd)		'�귣���ǥ�����ڵ�
						brnd_desc			= html2db(itemList.get(i).data.brnd_desc)			'�귣�弳��
						synonym_keyword		= html2db(itemList.get(i).data.synonym_keyword)		'�귣�嵿�Ǿ�

						strSql = ""
						strSql = strSql & " IF NOT EXISTS(SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_lotteon_Brandlist] WHERE brnd_id = '"&brnd_id&"')"
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_lotteon_Brandlist] "
						strSql = strSql & " (brnd_sct_cd, type_keyword, use_yn, mod_date, brnd_nm_kr, brnd_nm_etc1, brnd_nm_etc2, brnd_nm, brnd_nm_en, brnd_id, brnd_nm_main_cd, brnd_desc, synonym_keyword) "
						strSql = strSql & " VALUES ( "
						strSql = strSql & " '"&brnd_sct_cd&"', '"&type_keyword&"', '"&use_yn&"', '"&mod_date&"', '"&brnd_nm_kr&"', '"&brnd_nm_etc1&"', '"&brnd_nm_etc2&"', '"&brnd_nm&"', "
						strSql = strSql & "	'"&brnd_nm_en&"', '"&brnd_id&"', '"&brnd_nm_main_cd&"', '"&brnd_desc&"', '"&synonym_keyword&"' "
						strSql = strSql & " ) "
						dbget.Execute(strSql)

						' rw "�귣���������ڵ� : " & brnd_sct_cd
						' rw "�귣���Ÿ���� : " & type_keyword
						' rw "��뿩�� : " & use_yn
						' rw "�����Ͻ� : " & mod_date
						' rw "�귣���ѱ۸� : " & brnd_nm_kr
						' rw "�귣��ΰ�ù��°�̹���URL : " & brnd_logo_img_url1
						' rw "�귣���Ÿ��1 : " & brnd_nm_etc1
						' rw "�귣���Ÿ��2 : " & brnd_nm_etc2
						' rw "�귣���ǥ�� : " & brnd_nm
						' rw "�귣�念���� : " & brnd_nm_en
						' rw "�귣��ID : " & brnd_id
						' rw "�귣���ǥ�����ڵ� : " & brnd_nm_main_cd
						' rw "�귣�弳�� : " & brnd_desc
						' rw "�귣�嵿�Ǿ� : " & synonym_keyword
						' rw "---------------------------------------------------"
					Next
					rw "�귣�� " & skip & " ���� " & itemList.length & " �� ���"
				Set itemList = nothing
				Else
					rw "���̻� ����"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�����ڵ� �׷���ȸ
Public Function fnlotteonGetGroupCode()
    Dim objXML, iRbody, strObj, returnCode, datalist, i
	Dim grpCd, grpCdNm, grpCdEpn
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIURL & "/v1/openapi/bocommon/v1/code/getGroupCodeList", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�׷���ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.returnCode
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							grpCd		= datalist.get(i).grpCd			'�׷��ڵ�
							grpCdNm		= datalist.get(i).grpCdNm		'�׷��ڵ��
							grpCdEpn	= datalist.get(i).grpCdEpn		'�׷��ڵ弳��

							rw "�׷��ڵ� : " & grpCd
							rw "�׷��ڵ�� : " & grpCdNm
							rw "�׷��ڵ弳�� : " & grpCdEpn
							rw "--------------------------------------"
						Next
					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||����[�׷���ȸ] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'�����ڵ� �� ��ȸ
Public Function fnlotteonGetGroupCodeDetail(val)
    Dim objXML, iRbody, strObj, returnCode, datalist, i
	Dim grpCd, langCd, cd, cdNm, cdEpn, sortSeq, refcChrValEpn1, refcChrValEpn2, refcChrValEpn3, refcChrValEpn4
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIURL & "/v1/openapi/bocommon/v1/code/getDetailCodeList?grpCd="& val &"&langCd=ko", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�����ڵ����ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			' response.write iRbody
			' response.end

			Set strObj = JSON.parse(iRbody)
				returnCode		= strObj.returnCode
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							grpCd			= datalist.get(i).grpCd				'�׷��ڵ�
							langCd			= datalist.get(i).langCd			'����ڵ�
							cd				= datalist.get(i).cd				'�ڵ�
							cdNm			= datalist.get(i).cdNm				'�ڵ��
							cdEpn			= datalist.get(i).cdEpn				'�ڵ弳��
							sortSeq			= datalist.get(i).sortSeq			'���ļ���
							refcChrValEpn1	= datalist.get(i).refcChrValEpn1	'�������ڰ�����1
							refcChrValEpn2	= datalist.get(i).refcChrValEpn2	'�������ڰ�����2
							refcChrValEpn3	= datalist.get(i).refcChrValEpn3	'�������ڰ�����3
							refcChrValEpn4	= datalist.get(i).refcChrValEpn4	'�������ڰ�����4

							If val = "DV_CO_CD" Then		'�ù���ڵ�
								rw cd & "||" & cdNm
							Else
								rw "�׷��ڵ� : " & grpCd
								rw "����ڵ� : " & langCd
								rw "�ڵ� : " & cd
								rw "�ڵ�� : " & cdNm
								rw "�ڵ弳�� : " & cdEpn
								rw "���ļ��� : " & sortSeq
								rw "�������ڰ�����1 : " & refcChrValEpn1
								rw "�������ڰ�����2 : " & refcChrValEpn2
								rw "�������ڰ�����3 : " & refcChrValEpn3
								rw "�������ڰ�����4 : " & refcChrValEpn4
								rw "--------------------------------------"
							End If
						Next

					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||����[����ȸ] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'��ǰ ���
Public Function fnLotteonItemReg(iitemid, strParam, byRef iErrStr, imustprice, iLotteonSellYn, ilimityn, ilimitno, ilimiysold, iitemname, iimageNm, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, datalist, i, strSql
	Dim epdNo, spdNo, resultCode, resultMessage
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/product/registration/request", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ���] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							epdNo			= datalist.get(i).epdNo				'��ü��ǰ��ȣ
							spdNo			= datalist.get(i).spdNo				'�Ǹ��ڻ�ǰ��ȣ
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCrlf
						strSql = strSql & " SET lotteonregdate = getdate()" & VbCrlf
						If (spdNo <> "") Then
'							strSql = strSql & "	, lotteonStatCd = '3'"& VbCRLF					'���δ��
							strSql = strSql & "	, lotteonStatCd = '7'"& VbCRLF					'��ϿϷ�
						Else
							strSql = strSql & "	, lotteonStatCd = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & " ,lotteonGoodNo = '" & spdNo & "'" & VbCrlf
						strSql = strSql & " ,lotteonlastupdate = getdate()"
						strSql = strSql & " ,lotteonPrice = '"&imustprice&"' " & VbCrlf
						strSql = strSql & " ,lotteonsellYn = 'Y' "& VbCrlf
						strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
						strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
						strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_regitem R" & VbCrlf
						strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
						strSql = strSql & " where R.itemid = " & iitemid
						dbget.execute strSql
						iErrStr =  "OK||"&iitemid&"||��ϼ���(��ǰ���)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&resultMessage&"(Err1.��ǰ���)"
					End If
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.��ǰ���)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(��ǰ���)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ ����ȸ
Public Function fnLotteonItemView(iitemid, strParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, itmLst, i, strSql
	Dim epdNo, spdNo, resultCode, resultMessage
	Dim slStatCd, AssignedRow, fnlAprvYn, fnlAprvYnStr
	Dim outmallOptCode, outmallOptName, outmallSellyn, outmalllimitno, lotteonsellYn
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/product/detail", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ��ȸ] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					strSql = ""
					strSql =  strSql & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE mallid='"&CMALLNAME&"' and itemid="&iitemid&" "
					dbget.Execute strSql

					lotteonsellYn = strObj.data.slStatCd	'�ǸŻ����ڵ� [�����ڵ� : SL_STAT_CD] | END : �Ǹ�����, SALE : �Ǹ���, SOUT : ǰ��, STP : �Ǹ�����
					'rw strObj.data.slStatRsnCd				'��ǰ�ǸŻ��»����ڵ� [�����ڵ� : SL_STAT_RSN_CD] | SOUT_STK : ������ 0 ǰ��ó�� ,SOUT_LMT : �������� 0 ǰ��ó�� ,SOUT_ADMRSOUT_ADMR : ������ ����ǰ��ó�� ,SOUT_ITM : ����ǰǰ���� ��ǰ ǰ��ó�� ,SOUT_RSV : �����ǰ ǰ��ó�� ,SOUT_THDY : ���� ǰ������ ,STP_TNS : TNS(��Ģ��/�Ű�) �г�Ƽ �Ǹ����� ,STP_CTL : īŻ�α׿� ���� �Ǹ����� ,STP_MNTR : ��ǰ���� ����͸� ������ ó�� ,STP_DV : ��ۼ��� �г�Ƽ �Ǹ����� ,STP_UNAPRV : ��ǰ�̽��� ���¿� ���� �Ǹ�����ó�� ,END_TR : �ŷ�ó�������� ���� �Ǹ����� ó�� ,END_ADMR : ������ ���� �Ǹ����� ó�� ,END_TNS : TNS(���ػ�ǰ) �Ǹ����� ó��
					fnlAprvYn = strObj.data.pdAprvStatInfo.fnlAprvYn    '�������ο���
					fnlAprvYnStr = Chkiif(fnlAprvYn="Y", "���οϷ�", "������")
					Set itmLst = strObj.data.itmLst		'��ǰ���
						For i=0 to itmLst.length-1
							outmallSellyn = ""
							outmallOptCode	= itmLst.get(i).sitmNo			'�Ǹ��ڴ�ǰ��ȣ
							outmallOptName	= itmLst.get(i).sitmNm			'�Ǹ��ڴ�ǰ��
							If itmLst.get(i).slStatCd = "SALE" Then			'�ǸŻ����ڵ� [�����ڵ� : SL_STAT_CD] | SALE : �Ǹ���, SOUT : ǰ��
								outmallSellyn = "Y"
							Else
								outmallSellyn = "N"
							End If
'							rw itmLst.get(i).slPrc			'�ǸŰ�
							outmalllimitno	= itmLst.get(i).stkQty			'������ | ���������ΰ� Y�� ��쿡�� �ʼ���

							strSql = " INSERT INTO db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " VALUES ("&iitemid
							If i = 0 AND outmallOptName = "���ϻ�ǰ" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& i &"'" ''�ӽ÷� �Ե� �ڵ� ���� //2013/04/01
							End If
							strSql = strSql & " ,'"&CMALLNAME&"'"
							strSql = strSql & " ,'"&outmallOptCode&"'"
							strSql = strSql & " ,'"&html2DB(outmallOptName)&"'"
							strSql = strSql & " ,'"&outmallSellyn&"'"
							strSql = strSql & " ,'Y'"
							strSql = strSql & " ,"&outmalllimitno
							strSql = strSql & ")"
							dbget.Execute strSql, AssignedRow

							If (AssignedRow > 0) Then
								strSql = ""
								strSql = strSql & " UPDATE oP " & VbCRLF
								strSql = strSql & " SET itemoption = O.itemoption " & VbCRLF
								strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption oP " & VbCRLF
								strSql = strSql & " JOIN db_item.dbo.tbl_item_option o " & VbCRLF
								strSql = strSql & "     on oP.itemid = o.itemid " & VbCRLF
								strSql = strSql & " WHERE oP.mallid = '"&CMALLNAME&"' " & VbCRLF
								strSql = strSql & " and o.itemid = "&iitemid & VbCRLF
								strSql = strSql & " and oP.itemid = "&iitemid & VbCRLF
								strSql = strSql & " and op.outmallOptCode = '"&outmallOptCode&"'" & VbCRLF
								strSql = strSql & " and Replace(Replace(op.outmallOptName,' ',''),':','') = Replace(Replace(o.optionname,' ',''),':','')" & VbCRLF
								dbget.Execute strSql
							End If
						Next
					Set itmLst = nothing

					strSql = ""
					strSql = strSql & " UPDATE R " & VbCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VbCRLF
                    If fnlAprvYn = "Y" Then
                        strSql = strSql & " ,lotteonStatCd = 7"& VbCRLF
                    End If
					strSql = strSql & " ,lastStatcheckdate = getdate()"& VbCRLF
					strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_regItem R " & VbCRLF
					strSql = strSql & " JOIN ( " & VbCRLF
					strSql = strSql & " 	SELECT R.itemid,count(*) as CNT "
					strSql = strSql & " 	, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "		FROM db_etcmall.dbo.tbl_lotteon_regItem R " & VbCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro " & VbCRLF
					strSql = strSql & " 		on R.itemid = Ro.itemid"   & VbCRLF
					strSql = strSql & " 		and Ro.mallid = '"&CMALLNAME&"'"   & VbCRLF
					strSql = strSql & "         and Ro.itemid = "&iitemid & VbCRLF
					strSql = strSql & " 	GROUP BY R.itemid "   & VbCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VbCRLF
					dbget.Execute strSql
					iErrStr =  "OK||"&iitemid&"||����("&fnlAprvYnStr&")"
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.��ǰ��ȸ)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(��ǰ��ȸ)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CHKSTAT-002]"
		End If
	Set objXML= nothing
End Function

'���� ��ǰ ����
Public Function fnLotteonItemEdit(iitemid, iitemname, strParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, itmLst, i, strSql
	Dim epdNo, spdNo, resultCode, resultMessage
	Dim datalist, AssignedRow
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/product/modification/request", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set regitemname = '"&html2db(iitemname)&"'" & vbcrlf
						strSql = strSql & "	From db_etcmall.dbo.tbl_lotteon_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||����(��ǰ����)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& resultMessage &"����(��ǰ����)"
					End If

				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.��ǰ����)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(��ǰ����)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-002]"
		End If
	Set objXML= nothing
End Function

'���� ��ǰ ����..ȣ���� ���� ������ �ʹ� ������ �ϳ��� �����غ�..2020-05-07
Public Function fnLotteonItemEdit2(iitemid, iitemname, imustprice, strParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, itmLst, i, strSql
	Dim epdNo, spdNo, resultCode, resultMessage
	Dim datalist, AssignedRow
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/product/modification/request", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	SET regitemname = '"&html2db(iitemname)&"'" & vbcrlf
						strSql = strSql & "	, lotteonLastUpdate=getdate() " & vbcrlf
						strSql = strSql & "	, lotteonPrice = " & imustprice & VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	FROM db_etcmall.dbo.tbl_lotteon_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||����(��ǰ����)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& resultMessage &"����(��ǰ����)"
					End If

				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.��ǰ����)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(��ǰ����)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDIT-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ ��� ����
Public Function fnLotteOnQuantity(iitemid, istrParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, datalist, i, strSql
	Dim resultCode, resultMessage, failCnt, sitmNo
	On Error Resume Next
	failCnt = 0

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/item/stock/change", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[���] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						sitmNo = ""
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
							sitmNo			= datalist.get(i).sitmNo			'�Ǹ��ڴ�ǰ��ȣ
							If resultCode <> "0000" Then
								sitmNo = sitmNo & ","
								failCnt = failCnt + 1
							End If
						Next
						If Right(sitmNo,1) = "," Then
							sitmNo = Left(sitmNo, Len(sitmNo) - 1)
						End If
					Set datalist = nothing

					If failCnt = 0 Then
						iErrStr =  "OK||"&iitemid&"||��������(���)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"����(���)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-QTY-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ ���� ����
Public Function fnLotteOnPrice(iitemid, istrParam, imustprice, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, datalist, i, strSql
	Dim resultCode, resultMessage, failCnt, sitmNo
	On Error Resume Next
	failCnt = 0
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/item/price/change", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[����] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						sitmNo = ""
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
							sitmNo			= datalist.get(i).sitmNo			'�Ǹ��ڴ�ǰ��ȣ
							If resultCode <> "0000" Then
								sitmNo = sitmNo & ","
								failCnt = failCnt + 1
							End If
						Next
						If Right(sitmNo,1) = "," Then
							sitmNo = Left(sitmNo, Len(sitmNo) - 1)
						End If
					Set datalist = nothing

					If failCnt = 0 Then
						'// ��ǰ�������� ����
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem  " & VbCRLF
						strSql = strSql & "	SET lotteonLastUpdate=getdate() " & VbCRLF
						strSql = strSql & "	, lotteonPrice = " & imustprice & VbCRLF
						strSql = strSql & "	,accFailCnt = 0"& VbCRLF
						strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"����(��ǰ����)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICE-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ �ǸŻ��� ����
Public Function fnLotteOnSellyn(iitemid, ichgSellYn, strParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, datalist, i, strSql
	Dim resultCode, resultMessage
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/product/status/change", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�ǸŻ���] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						If ichgSellyn = "Y" Then
							strSql = ""
							strSql = strSql & " UPDATE R"
							strSql = strSql & "	Set lotteonSellYn = 'Y'"
							strSql = strSql & "	,lotteonLastUpdate = getdate()"
							strSql = strSql & "	From db_etcmall.dbo.tbl_lotteon_regitem  R"
							strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
							dbget.Execute(strSql)
							iErrStr =  "OK||"&iitemid&"||�Ǹ�(���º���)"
						ElseIf ichgSellyn = "N" Then
							strSql = ""
							strSql = strSql & " UPDATE R"
							strSql = strSql & "	Set lotteonSellYn = 'N'"
							strSql = strSql & "	,accFailCnt = 0"
							strSql = strSql & "	,lotteonLastUpdate = getdate()"
							strSql = strSql & "	From db_etcmall.dbo.tbl_lotteon_regitem  R"
							strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
							dbget.Execute(strSql)
							iErrStr =  "OK||"&iitemid&"||ǰ��ó��(���º���)"
						ElseIf ichgSellYn = "X" Then
							strSql = ""
							strSql = strSql &" INSERT INTO [db_etcmall].[dbo].[tbl_Outmall_Delete_Log] " & VBCRLF
							strSql = strSql &" SELECT TOP 1 'lotteon', i.itemid, r.lotteonGoodNo, r.lotteonRegdate, getdate(), r.lastErrStr" & VBCRLF
							strSql = strSql &" FROM db_item.dbo.tbl_item as i " & VBCRLF
							strSql = strSql &" JOIN db_etcmall.dbo.tbl_lotteon_regitem as r on i.itemid = r.itemid " & VBCRLF
							strSql = strSql &" WHERE i.itemid = "&iitemid & VBCRLF
							dbget.Execute(strSql)

							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_lotteon_regitem " & vbcrlf
							strSql = strSql & " WHERE itemid = '"&iitemid&"' "
							dbget.Execute(strSql)

							strSql = ""
							strSql = strSql & " DELETE FROM db_item.dbo.tbl_outmall_regedoption " & vbcrlf
							strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
							strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
							dbget.Execute(strSql)

							strSql = ""
							strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que " & vbcrlf
							strSql = strSql & " WHERE itemid = '"&iitemid&"' " & vbcrlf
							strSql = strSql & " and mallid = '"&CMALLNAME&"' " & vbcrlf
							dbget.Execute(strSql)
							iErrStr = "OK||"&iitemid&"||�Ǹ�����"
						End If
					Else
						If InStr(resultMessage, "�ߺ���ǰ���� ���� �Ǹ�������") Then
							strSql = ""
							strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regItem " & VbCrlf
							strSql = strSql & " SET lotteonlastupdate = getdate()" & VbCrlf
							strSql = strSql & " ,accFailCNT=0" & VbCrlf
							strSql = strSql & " ,lotteonSellYn = 'N'" & VbCRLF
							strSql = strSql & " WHERE itemid = " & iitemid
							dbget.execute strSql

							strSql = ""
							strSql = strSql & " IF NOT Exists(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE itemid='"&iitemid&"' and mallgubun = '"&CMALLNAME&"') "
							strSql = strSql & "  BEGIN "
							strSql = strSql & "  	INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid, mallgubun, bigo) VALUES('"&iitemid&"','"&CMALLNAME&"', '�������� �ʰų� �̹��� ����') "
							strSql = strSql & "  END "
							dbget.Execute strSql
							iErrStr = "OK||"&iitemid&"||�Ǹ�����(���º���)/������ ����ó��"
						Else
							iErrStr = "ERR||"&iitemid&"||"&resultMessage&"(Err1.�ǸŻ���)"
						End If
					End If
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.�ǸŻ���)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(�ǸŻ���)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-EDITSELLYN-002]"
		End If
	Set objXML= nothing
End Function

'��ǰ �ǸŻ��� ����
Public Function fnLotteOnOptStat(iitemid, strParam, iErrStr, iresponseJson)
    Dim objXML, iRbody, strObj, returnCode, resultmsg, datalist, i, strSql
	Dim resultCode, resultMessage, failCnt, sitmNo
	On Error Resume Next
	failCnt = 0

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/product/v1/item/status/change", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ǰ����] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			If iresponseJson = "Y" Then
				response.write iRbody
			End If
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				resultmsg	= replaceMsg(strObj.message)
				If returnCode = "0000" Then
					Set datalist = strObj.data
						sitmNo = ""
						For i=0 to datalist.length-1
							resultCode		= datalist.get(i).resultCode		'ó�����
							resultMessage	= datalist.get(i).resultMessage		'ó���޼���
							sitmNo			= datalist.get(i).sitmNo			'�Ǹ��ڴ�ǰ��ȣ
							If resultCode <> "0000" Then
								sitmNo = sitmNo & ","
								failCnt = failCnt + 1
							End If
						Next
						If Right(sitmNo,1) = "," Then
							sitmNo = Left(sitmNo, Len(sitmNo) - 1)
						End If
					Set datalist = nothing

					If failCnt = 0 Then
						iErrStr =  "OK||"&iitemid&"||��������(��ǰ����)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"����(��ǰ����)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSTAT-002]"
		End If
	Set objXML= nothing
End Function

'���� ó�� �Ϸ�
Function fnlotteonTenConfirmOrder(vOdNo, vodSeq, vProcSeq)
	Dim objXML, xmlDOM, iRbody, strSql
	Dim obj, strParam, strObj
	Dim vDvTrcStatDttm, returnCode
	vDvTrcStatDttm = FormatDate(now(), "00000000000000")

	Set obj = jsObject()
		Set obj("ifCompleteList")= jsArray()									'����ó���Ϸ���
			Set obj("ifCompleteList")(null) = jsObject()
				obj("ifCompleteList")(null)("dvRtrvDvsCd") = "DV"					'#���ȸ�������ڵ� | DV:���, RTRV:ȸ��
				obj("ifCompleteList")(null)("odNo") = vOdNo							'#�ֹ���ȣ : �ֹ����̺��� PK�Ӽ�
				obj("ifCompleteList")(null)("odSeq") = vodSeq						'#�ֹ����� : �ֹ������� ���ؼ� ��ǰ���� �ο��Ǵ� �Ӽ��� 1
				obj("ifCompleteList")(null)("procSeq") = vProcSeq					'#ó������ : Default 1 ��ǰ������ ó���������� ������. ���� �Է½� 1 �̰� Ŭ������ �߻��� ��� 1�� ������
				obj("ifCompleteList")(null)("orglProcSeq") = ""						'oŬ������ ��� �ʼ� ��ó������ : Ŭ������ �߻����� ��� ����ó������
				obj("ifCompleteList")(null)("clmNo") = ""							'oŬ������ ��� �ʼ� Ŭ���ӹ�ȣ : Ŭ������ �߻����� ����� ��ȣ
				obj("ifCompleteList")(null)("ifCplYN") = "Y"						'#�����ϷῩ��(Y/N)
				obj("ifCompleteList")(null)("ifFlRsnCnts") = ""						'#�������� ��������
		 		strParam = obj.jsString
	Set obj = nothing

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerIfCompleteInform", false				'���/ȸ������ �����Ϸ� �뺸
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[����Ȯ��] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				If returnCode = "0000" Then
					If strObj.data.rsltCd = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
						strSql = strSql & " SET isTenConfirmSend = 'Y' "
						strSql = strSql & " WHERE outmallorderserial = '"&vOdNo&"'  "
						strSql = strSql & " and beasongNum11st = '"&vProcSeq&"' "
						strSql = strSql & " and OrgDetailKey = '"&vodSeq&"' "
						strSql = strSql & " and mallid = 'lotteon' "
						dbget.Execute strSql
						fnlotteonTenConfirmOrder = true
					Else
						rw vOdNo & " ���� : " & strObj.data.rsltMsg
					End If
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

Function fnlotteonConfirmOrder(vOdNo, vodSeq, vProcSeq, vSpdNo, vSitmNo, vSlQty)
	Dim objXML, xmlDOM, iRbody, strSql
	Dim obj, strParam, strObj
	Dim vDvTrcStatDttm, returnCode
	vDvTrcStatDttm = FormatDate(now(), "00000000000000")

	Set obj = jsObject()
		Set obj("deliveryProgressStateList")= jsArray()									'��ۻ��¸��
			Set obj("deliveryProgressStateList")(null) = jsObject()
				obj("deliveryProgressStateList")(null)("odNo") = vOdNo						'#�ֹ���ȣ : �ֹ����̺��� PK�Ӽ�
				obj("deliveryProgressStateList")(null)("odSeq") = vodSeq					'#�ֹ����� : �ֹ������� ���ؼ� ��ǰ���� �ο��Ǵ� �Ӽ��� 1
				obj("deliveryProgressStateList")(null)("procSeq") = vProcSeq				'#ó������ : Default 1 ��ǰ������ ó���������� ������. ���� �Է½� 1 �̰� Ŭ������ �߻��� ��� 1�� ������
				obj("deliveryProgressStateList")(null)("odPrgsStepCd") = "12"				'#�ֹ�����ܰ� | 11 : �������, 12 : ��ǰ�غ�, 13 : �߼ۿϷ�, 14 : ��ۿϷ�, 15 : ����Ϸ�, 23 : ȸ������, 24 : ȸ������, 25 : ȸ���Ϸ�, 26 : ȸ��Ȯ��
				obj("deliveryProgressStateList")(null)("dvTrcStatDttm") = vDvTrcStatDttm	'#��ۻ��¹߻��Ͻ�
				obj("deliveryProgressStateList")(null)("spdNo") = vSpdNo					'#��ǰ��ȣ : �Ե�ON���� �����Ǵ� ��ǰ��ȣ
				obj("deliveryProgressStateList")(null)("sitmNo") = vSitmNo					'#��ǰ��ȣ : �Ե�ON���� �����Ǵ� ��ǰ��ȣ
				obj("deliveryProgressStateList")(null)("slQty") = vSlQty					'#���� : ��ǰ�� ���� �ֹ�����
		 		strParam = obj.jsString
	Set obj = nothing

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerDeliveryProgressStateInform", false		'��ۻ����뺸
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[�ֹ�Ȯ��] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				If returnCode = "0000" Then
					If strObj.data.rsltCd = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
						strSql = strSql & " SET isbaljuConfirmSend = 'Y' "
						strSql = strSql & " , lastUpdate = getdate() "
						strSql = strSql & " WHERE outmallorderserial = '"&vOdNo&"'  "
						strSql = strSql & " and beasongNum11st = '"&vProcSeq&"' "
						strSql = strSql & " and OrgDetailKey = '"&vodSeq&"' "
						strSql = strSql & " and mallid = 'lotteon' "
						dbget.Execute strSql
						fnlotteonConfirmOrder = true
					Else
						rw vOdNo & " ���� : " & strObj.data.rsltMsg
					End If
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'�Ե�ON��ۻ��� ��ȸ
Function fnlotteonViewOrder(vOrderNo)
	Dim objXML, xmlDOM, iRbody, strSql
	Dim obj, strParam, strObj, searchDate
	Dim returnCode

	strSql = ""
	strSql = strSql & " SELECT TOP 1 deliveryDate "
	strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMP11stOrder "
	strSql = strSql & " WHERE mallid = 'lotteon' "
	strSql = strSql & " and outmallorderserial = '"& vOrderNo &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		searchDate = Replace(LEFT(rsget("deliveryDate"), 10), "-", "")
	End If
	rsget.Close

	Set obj = jsObject()
		obj("srchStrtDt") = searchDate & "000000"			'#�˻��������� yyyymmddhhmmss ������û����Ͻ�
		obj("srchEndDt") =  searchDate & "235959"			'#�˻��������� yyyymmddhhmmss
		obj("odNo") = vOrderNo			'�ֹ���ȣ
		strParam = obj.jsString
	Set obj = nothing

	rw strParam

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerDeliveryProgressStateSearch", false
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||����[��ۻ���] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			response.write iRbody
			response.end

			Set strObj = JSON.parse(iRbody)
				returnCode	= strObj.returnCode
				If returnCode = "0000" Then
					rw strObj.data.odPrgsStepCd
				Else
					rw strObj.message
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function
'############################################## ���� �����ϴ� API �Լ� ���� �� ############################################
%>
