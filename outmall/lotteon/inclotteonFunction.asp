<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'판매자 출고지/반품지 리스트조회
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
			iErrStr = "ERR||"&iitemid&"||실패[DVP] " & html2db(Err.Description)
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
							dvpNo			= datalist.get(i).dvpNo			''PLE&PLE'+배송지번호
							dvpTypCd		= datalist.get(i).dvpTypCd		'배송지유형코드 [회수지:01, 출고지:02]
							dvpNm			= datalist.get(i).dvpNm			'배송지명
							zipNo			= datalist.get(i).zipNo			'우편번호
							zipAddr			= datalist.get(i).zipAddr		'우편주소
							dtlAddr			= datalist.get(i).dtlAddr		'상세주소
							stnmZipNo		= datalist.get(i).stnmZipNo		'도로명우편번호
							stnmZipAddr		= datalist.get(i).stnmZipAddr	'도로명우편주소
							stnmDtlAddr		= datalist.get(i).stnmDtlAddr	'도루명상세주소
							rpbtrNm			= datalist.get(i).rpbtrNm		'담당자명
							mphnNatnNoCd	= datalist.get(i).mphnNatnNoCd	'휴대폰국가코드 [default:'82']
							mphnNo			= datalist.get(i).mphnNo		'휴대폰번호
							telNatnNoCd		= datalist.get(i).telNatnNoCd	'연락처국가코드 [defaul:'82']
							telNo			= datalist.get(i).telNo			'연락처
							lrtrNo			= datalist.get(i).lrtrNo		'??하위거래처번호
							useYn			= datalist.get(i).useYn			'사용여부 [default:'Y']

							If useYn = "Y" Then
								rw "'PLE&PLE'+배송지번호 : " & dvpNo
								rw "배송지유형코드 [회수지:01, 출고지:02] : " & dvpTypCd
								rw "배송지명 : " & dvpNm
								rw "우편번호 : " & zipNo
								rw "우편주소 : " & zipAddr
								rw "상세주소 : " & dtlAddr
								rw "도로명우편번호 : " & stnmZipNo
								rw "도로명우편주소 : " & stnmZipAddr
								rw "도루명상세주소 : " & stnmDtlAddr
								rw "담당자명 : " & rpbtrNm
								rw "휴대폰국가코드 : " & mphnNatnNoCd
								rw "휴대폰번호 : " & mphnNo
								rw "연락처국가코드 : " & telNatnNoCd
								rw "연락처 : " & telNo
								rw "??하위거래처번호 : " & lrtrNo
								rw "사용여부 : " & useYn
								rw "--------------------------------------"
							End If
						Next
					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||실패[그룹조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'속성 기본 조회
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
			iErrStr = "ERR||"&iitemid&"||실패[그룹조회] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList
				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						attr_id			= ""
						attr_id			= itemList.get(i).data.attr_id					'속성유형 ID
						attr_val_frm_cd = itemList.get(i).data.attr_val_frm_cd			'속성값형식구분코드
						use_yn			= itemList.get(i).data.use_yn					'사용여부
						mod_date		= itemList.get(i).data.mod_date					'수정일시
						attr_nm			= itemList.get(i).data.attr_nm					'속성유형명
						attr_pi_type	= itemList.get(i).data.attr_pi_type				'속성상품형태구분
						attr_disp_nm	= itemList.get(i).data.attr_disp_nm				'속성유형전시명

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_Attribute_Ins] '"&attr_id&"', '"&attr_val_frm_cd&"', '"&use_yn&"', '"&mod_date&"' " & VBCRLF
						strSql = strSql & " ,'"&attr_nm&"' ,'"&attr_pi_type&"' ,'"&attr_disp_nm&"'  "
   						dbget.Execute(strSql)

						' rw "속성유형 ID : " & attr_id
						' rw "속성값형식구분코드 : " & attr_val_frm_cd
						' rw "사용여부 : " & use_yn
						' rw "수정일시 : " & mod_date
						' rw "속성유형명 : " & attr_nm
						' rw "속성상품형태구분 : " & attr_pi_type
						' rw "속성유형전시명 : " & attr_disp_nm
						' rw "--------------------------------------"
						Set attr_val_list = itemList.get(i).data.attr_val_list
						 	For j = 0 to attr_val_list.length-1
								attr_ref_val1		= ""
								attr_ref_val2		= ""
								attr_val_nm			= ""
								attr_val_disp_nm	= ""
								use_yn2				= ""
								attr_val_id			= ""

								attr_ref_val1		= attr_val_list.get(j).attr_ref_val1		'속성참조값1 (색상의 경우 색상계열, 사이즈의 경우 상위계열)
								attr_ref_val2		= attr_val_list.get(j).attr_ref_val2		'속성참조값2 (사이즈의 경우 같은 레벨로 묶이는 정보)
								attr_val_nm			= attr_val_list.get(j).attr_val_nm			'속성값명
								attr_val_disp_nm	= attr_val_list.get(j).attr_val_disp_nm		'속성값전시명
								use_yn2				= attr_val_list.get(j).use_yn				'사용여부
								attr_val_id			= attr_val_list.get(j).attr_val_id			'속성값ID

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_AttributeValue_Ins] '"&attr_id&"', '"&attr_val_id&"', '"&attr_val_nm&"' " & VBCRLF
								strSql = strSql & " ,'"&attr_val_disp_nm&"' ,'"&attr_ref_val1&"' ,'"&attr_ref_val2&"','"&use_yn2&"'  "
								dbget.Execute(strSql)
								' rw "속성참조값1 : " & attr_ref_val1
								' rw "속성참조값2 : " & attr_ref_val2
								' rw "속성값명 : " & attr_val_nm
								' rw "속성값전시명 : " & attr_val_disp_nm
								' rw "사용여부 : " & use_yn2
								' rw "속성값ID : " & attr_val_id
						 	Next
						Set attr_val_list = nothing
						' rw "###################################################"
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=ATTRVIEW&rSkip=0
						' rw "###################################################"
					Next
					rw "속성 " & skip & " 부터 " & itemList.length & " 건 등록"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					rw "더이상 없음"
					ihasNext = "N"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'전시카테고리 조회
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

		rw "전시카테고리 생성 완료"
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
			iErrStr = "ERR||"&iitemid&"||실패[그룹조회] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList

				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						disp_cat_id		= itemList.get(i).data.disp_cat_id			'전시카테고리ID
						disp_cat_nm 	= itemList.get(i).data.disp_cat_nm			'전시카테고리명
						disp_cat_desc	= itemList.get(i).data.disp_cat_desc		'전시카테고리설명
						upr_disp_cat_id	= itemList.get(i).data.upr_disp_cat_id		'상위전시카테고리번호
						depth_no		= itemList.get(i).data.depth_no				'깊이번호
						leaf_yn			= itemList.get(i).data.leaf_yn				'리프여부
						prio_rnk		= itemList.get(i).data.prio_rnk				'우선순위 (전시카테고리 우선순위)
						mall_dvs_cd		= itemList.get(i).data.mall_dvs_cd			'몰구분코드
						disp_yn			= itemList.get(i).data.disp_yn				'전시여부
						mod_date		= itemList.get(i).data.mod_date				'수정일시
						use_yn			= itemList.get(i).data.use_yn				'사용여부

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_TmpDispCategory_Ins] '"&disp_cat_id&"', '"&disp_cat_nm&"', '"&disp_cat_desc&"', '"&upr_disp_cat_id&"' " & VBCRLF
						strSql = strSql & " ,'"&depth_no&"' ,'"&leaf_yn&"' ,'"&prio_rnk&"' ,'"&mall_dvs_cd&"' ,'"&disp_yn&"' ,'"&mod_date&"' ,'"&use_yn&"'  "
   						dbget.Execute(strSql)

						'전시카테고리 개편시 사용법
						'1. http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=DISPCATE&rSkip=0
						'2. 아래 Url 호출하여 카테고리 생성
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=DISPCATE&rSkip=skip

						' rw "전시카테고리ID : " & disp_cat_id
						' rw "전시카테고리명 : " & disp_cat_nm
						' rw "전시카테고리설명 : " & disp_cat_desc
						' rw "상위전시카테고리번호 : " & upr_disp_cat_id
						' rw "깊이번호 : " & depth_no
						' rw "리프여부 : " & leaf_yn
						' rw "우선순위 : " & prio_rnk
						' rw "몰구분코드 : " & mall_dvs_cd
						' rw "전시여부 : " & disp_yn
						' rw "수정일시 : " & mod_date
						' rw "사용여부 : " & use_yn
						' rw "---------------------------------------------------"
					Next
					rw "전시카테고리 " & skip & " 부터 " & itemList.length & " 건 등록"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					rw "더이상 없음"
					ihasNext = "N"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'표준카테고리 조회
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

		rw "표준카테고리 생성 완료"
		response.end
		Exit Function
	End If

'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", APIATTRURL & "/cheetah/econCheetah.ecn?job=cheetahStandardCategory&skip="&skip&"&limit=500&sort=sort_2&direction=asc", false		'sort_1 = 수정일 / sort_2 = 전시카테고리ID
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[표준조회] " & html2db(Err.Description)
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
						std_cat_id		= itemList.get(i).data.std_cat_id			'표준카테고리ID
'						rw "표준카테고리ID : " & std_cat_id

						Set disp_list = itemList.get(i).data.disp_list				'전시카테고리 리스트
							For j = 0 to disp_list.length-1
								dstd_cat_id 	= ""
								dmall_dvs_cd 	= ""
								ddisp_cat_id 	= ""

								dstd_cat_id		= disp_list.get(j).std_cat_id			'표준카테고리ID
								dmall_dvs_cd	= disp_list.get(j).mall_dvs_cd			'몰구분코드
								ddisp_cat_id	= disp_list.get(j).disp_cat_id			'전시카테고리ID

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_StdCategoryDisp_Ins] '"&dstd_cat_id&"', '"&ddisp_cat_id&"', '"&dmall_dvs_cd&"' "
								dbget.Execute(strSql)
'								rw " l 표준카테고리ID : " & dstd_cat_id
'								rw " l 몰구분코드 : " & dmall_dvs_cd
'								rw " l 전시카테고리ID : " & ddisp_cat_id
'								rw "lllllllllllllllllllllllllllllllllllllllllllllllllll"
							Next
						Set disp_list = nothing

						Set attr_list = itemList.get(i).data.attr_list				'속성유형 리스트
							For k = 0 to attr_list.length-1
								astd_cat_id	= ""
								aattr_id	= ""
								aprio_rnk	= ""

								astd_cat_id	= attr_list.get(k).std_cat_id				'표준카테고리ID
								aattr_id	= attr_list.get(k).attr_id					'속성유형ID
								aprio_rnk	= attr_list.get(k).prio_rnk					'우선순위 (표준카테고리 - 속성유형 맵핑에 대한 우선순위)

								strSql = ""
								strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_StdCategoryAttr_Ins] '"&astd_cat_id&"', '"&aattr_id&"', '"&aprio_rnk&"' "
								dbget.Execute(strSql)
'								rw " a 표준카테고리ID : " & astd_cat_id
'								rw " a 속성유형ID : " & aattr_id
'								rw " a 우선순위 : " & aprio_rnk
'								rw "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
							Next
						Set attr_list = nothing
'						Set pd_itms_list = itemList.get(i).data.pd_itms_list		'상품품목고시 정보 리스트
						std_cat_nm		= itemList.get(i).data.std_cat_nm			'표준카테고리명
						std_cat_desc	= itemList.get(i).data.std_cat_desc			'표준카테고리설명
						upr_std_cat_id	= itemList.get(i).data.upr_std_cat_id		'상위표준카테고리번호
						depth_no		= itemList.get(i).data.depth_no				'깊이번호
						leaf_yn			= itemList.get(i).data.leaf_yn				'리프여부
						prio_rnk		= itemList.get(i).data.prio_rnk				'우선순위 (표준카테고리 우선순위)
						tdf_cd			= itemList.get(i).data.tdf_cd				'과세구분코드
						age_limit_cd	= itemList.get(i).data.age_limit_cd			'나이제한코드
						deduct_cult_yn	= itemList.get(i).data.deduct_cult_yn		'도서문화비 공제여부
						aband_pickup_yn	= itemList.get(i).data.aband_pickup_yn		'폐가전 수거여부
						exch_money_yn	= itemList.get(i).data.exch_money_yn		'환금성 여부
						counsel_prod_yn	= itemList.get(i).data.counsel_prod_yn		'상담접수 상품여부
						mod_date		= itemList.get(i).data.mod_date				'수정일시
						use_yn			= itemList.get(i).data.use_yn				'사용여부
						chl_athn		= itemList.get(i).data.chl_athn				'[어린이제품]안전인증
						chl_cfm			= itemList.get(i).data.chl_cfm				'[어린이제품]안전확인
						chl_sups		= itemList.get(i).data.chl_sups				'[어린이제품]공급자적합성확인
						cmcn_athn		= itemList.get(i).data.cmcn_athn			'[방송통신기자재]적합인증
						cmcn_reg		= itemList.get(i).data.cmcn_reg				'[방송통신기자재]적합등록
						cmcn_tntt		= itemList.get(i).data.cmcn_tntt			'[방송통신기자재]잠정인증
						elc_athn		= itemList.get(i).data.elc_athn				'[전기용품]안전인증
						elc_cfm			= itemList.get(i).data.elc_cfm				'[전기용품]안전확인
						elc_sups		= itemList.get(i).data.elc_sups				'[전기용품]공급자적합성확인
						life_athn		= itemList.get(i).data.life_athn			'[생활용품]안전인증
						life_cfm		= itemList.get(i).data.life_cfm				'[생활용품]안전확인
						life_sups		= itemList.get(i).data.life_sups			'[생활용품]공급자적합성확인
						life_std		= itemList.get(i).data.life_std				'[생활용품]안전기준준수
						chem_life		= itemList.get(i).data.chem_life			'[화학제품] 생활화학제품 안전기준적합확인신고번호 / 승인번호
						chem_bioc		= itemList.get(i).data.chem_bioc			'[화학제품] 살생물제품 승인번호
						etc				= itemList.get(i).data.etc					'기타

						strSql = ""
						strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_Lotteon_TmpStdCategory_Ins] '"&std_cat_id&"', '"&std_cat_nm&"', '"&std_cat_desc&"', '"&upr_std_cat_id&"' " & VBCRLF
						strSql = strSql & " ,'"&depth_no&"' ,'"&leaf_yn&"' ,'"&prio_rnk&"' ,'"&tdf_cd&"' ,'"&age_limit_cd&"' ,'"&deduct_cult_yn&"' ,'"&aband_pickup_yn&"' ,'"&exch_money_yn&"' " & VBCRLF
						strSql = strSql & " ,'"&counsel_prod_yn&"' ,'"&mod_date&"' ,'"&use_yn&"' ,'"&chl_athn&"' ,'"&chl_cfm&"' ,'"&chl_sups&"' ,'"&cmcn_athn&"' ,'"&cmcn_reg&"' " & VBCRLF
						strSql = strSql & " ,'"&cmcn_tntt&"' ,'"&elc_athn&"' ,'"&elc_cfm&"' ,'"&elc_sups&"' ,'"&life_athn&"' ,'"&life_cfm&"' ,'"&life_sups&"' ,'"&life_std&"' " & VBCRLF
						strSql = strSql & " ,'"&chem_life&"' ,'"&chem_bioc&"' ,'"&etc&"' "
   						dbget.Execute(strSql)

						'표준카테고리 개편시 사용법
						'1.	http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=STDCATE&rSkip=0

						'2. 아래 Url 호출하여 카테고리 생성
						'http://localhost:11117/outmall/lotteon/actlotteonReq.asp?cmdparam=STDCATE&rSkip=skip

						' rw "표준카테고리명 : " & std_cat_nm
						' rw "표준카테고리설명 : " & std_cat_desc
						' rw "상위표준카테고리번호 : " & upr_std_cat_id
						' rw "깊이번호 : " & depth_no
						' rw "리프여부 : " & leaf_yn
						' rw "우선순위 : " & prio_rnk
						' rw "과세구분코드 : " & tdf_cd
						' rw "나이제한코드 : " & age_limit_cd
						' rw "도서문화비 공제여부 : " & deduct_cult_yn
						' rw "폐가전 수거여부 : " & aband_pickup_yn
						' rw "환금성 여부 : " & exch_money_yn
						' rw "상담접수 상품여부 : " & counsel_prod_yn
						' rw "수정일시 : " & mod_date
						' rw "사용여부 : " & use_yn
						' rw "[어린이제품]안전인증 : " & chl_athn
						' rw "[어린이제품]안전확인 : " & chl_cfm
						' rw "[어린이제품]공급자적합성확인 : " & chl_sups
						' rw "[방송통신기자재]적합인증 : " & cmcn_athn
						' rw "[방송통신기자재]적합등록 : " & cmcn_reg
						' rw "[방송통신기자재]잠정인증 : " & cmcn_tntt
						' rw "[전기용품]안전인증 : " & elc_athn
						' rw "[전기용품]안전확인 : " & elc_cfm
						' rw "[전기용품]공급자적합성확인 : " & elc_sups
						' rw "[생활용품]안전인증 : " & life_athn
						' rw "[생활용품]안전확인 : " & life_cfm
						' rw "[생활용품]공급자적합성확인 : " & life_sups
						' rw "[생활용품]안전기준준수 : " & life_std
						' rw "[화학제품] 생활화학제품 안전기준적합확인신고번호 / 승인번호 : " & chem_life
						' rw "[화학제품] 살생물제품 승인번호 : " & chem_bioc
						' rw "기타 : " & etc
						' rw "###################################################"
						' rw "###################################################"
					Next
					rw "표준카테고리 " & skip & " 부터 " & itemList.length & " 건 등록"
					ihasNext = "Y"
					skip = skip + 500
				Set itemList = nothing
				Else
					ihasNext = "N"
					rw "더이상 없음"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'브랜드 조회
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
			iErrStr = "ERR||"&iitemid&"||실패[브랜드조회] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				Set itemList = strObj.itemList

				If itemList.length > 0 Then
					For i=0 to itemList.length-1
						brnd_sct_cd			= html2db(itemList.get(i).data.brnd_sct_cd)			'브랜드유형명코드 (상품연계시: 상품브랜드)
						type_keyword		= html2db(itemList.get(i).data.typo_keyword)		'브랜드오타어목록
						use_yn				= itemList.get(i).data.use_yn						'사용여부
						mod_date			= itemList.get(i).data.mod_date						'수정일시
						brnd_nm_kr			= html2db(itemList.get(i).data.brnd_nm_kr)			'브랜드한글명
						brnd_logo_img_url1	= itemList.get(i).data.brnd_logo_img_url1			'브랜드로고첫번째이미지URL
						brnd_nm_etc1		= html2db(itemList.get(i).data.brnd_nm_etc1)		'브랜드기타명1
						brnd_nm_etc2		= html2db(itemList.get(i).data.brnd_nm_etc2)		'브랜드기타명2
						brnd_nm				= html2db(itemList.get(i).data.brnd_nm)				'브랜드대표명
						brnd_nm_en			= html2db(itemList.get(i).data.brnd_nm_en)			'브랜드영문명
						brnd_id				= html2db(itemList.get(i).data.brnd_id)				'브랜드ID
						brnd_nm_main_cd		= html2db(itemList.get(i).data.brnd_nm_main_cd)		'브랜드대표명구분코드
						brnd_desc			= html2db(itemList.get(i).data.brnd_desc)			'브랜드설명
						synonym_keyword		= html2db(itemList.get(i).data.synonym_keyword)		'브랜드동의어

						strSql = ""
						strSql = strSql & " IF NOT EXISTS(SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_lotteon_Brandlist] WHERE brnd_id = '"&brnd_id&"')"
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_lotteon_Brandlist] "
						strSql = strSql & " (brnd_sct_cd, type_keyword, use_yn, mod_date, brnd_nm_kr, brnd_nm_etc1, brnd_nm_etc2, brnd_nm, brnd_nm_en, brnd_id, brnd_nm_main_cd, brnd_desc, synonym_keyword) "
						strSql = strSql & " VALUES ( "
						strSql = strSql & " '"&brnd_sct_cd&"', '"&type_keyword&"', '"&use_yn&"', '"&mod_date&"', '"&brnd_nm_kr&"', '"&brnd_nm_etc1&"', '"&brnd_nm_etc2&"', '"&brnd_nm&"', "
						strSql = strSql & "	'"&brnd_nm_en&"', '"&brnd_id&"', '"&brnd_nm_main_cd&"', '"&brnd_desc&"', '"&synonym_keyword&"' "
						strSql = strSql & " ) "
						dbget.Execute(strSql)

						' rw "브랜드유형명코드 : " & brnd_sct_cd
						' rw "브랜드오타어목록 : " & type_keyword
						' rw "사용여부 : " & use_yn
						' rw "수정일시 : " & mod_date
						' rw "브랜드한글명 : " & brnd_nm_kr
						' rw "브랜드로고첫번째이미지URL : " & brnd_logo_img_url1
						' rw "브랜드기타명1 : " & brnd_nm_etc1
						' rw "브랜드기타명2 : " & brnd_nm_etc2
						' rw "브랜드대표명 : " & brnd_nm
						' rw "브랜드영문명 : " & brnd_nm_en
						' rw "브랜드ID : " & brnd_id
						' rw "브랜드대표명구분코드 : " & brnd_nm_main_cd
						' rw "브랜드설명 : " & brnd_desc
						' rw "브랜드동의어 : " & synonym_keyword
						' rw "---------------------------------------------------"
					Next
					rw "브랜드 " & skip & " 부터 " & itemList.length & " 건 등록"
				Set itemList = nothing
				Else
					rw "더이상 없음"
				End IF
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'공통코드 그룹조회
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
			iErrStr = "ERR||"&iitemid&"||실패[그룹조회] " & html2db(Err.Description)
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
							grpCd		= datalist.get(i).grpCd			'그룹코드
							grpCdNm		= datalist.get(i).grpCdNm		'그룹코드명
							grpCdEpn	= datalist.get(i).grpCdEpn		'그룹코드설명

							rw "그룹코드 : " & grpCd
							rw "그룹코드명 : " & grpCdNm
							rw "그룹코드설명 : " & grpCdEpn
							rw "--------------------------------------"
						Next
					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||실패[그룹조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'공통코드 상세 조회
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
			iErrStr = "ERR||"&iitemid&"||실패[공통코드상세조회] " & html2db(Err.Description)
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
							grpCd			= datalist.get(i).grpCd				'그룹코드
							langCd			= datalist.get(i).langCd			'언어코드
							cd				= datalist.get(i).cd				'코드
							cdNm			= datalist.get(i).cdNm				'코드명
							cdEpn			= datalist.get(i).cdEpn				'코드설명
							sortSeq			= datalist.get(i).sortSeq			'정렬순서
							refcChrValEpn1	= datalist.get(i).refcChrValEpn1	'참조문자값설명1
							refcChrValEpn2	= datalist.get(i).refcChrValEpn2	'참조문자값설명2
							refcChrValEpn3	= datalist.get(i).refcChrValEpn3	'참조문자값설명3
							refcChrValEpn4	= datalist.get(i).refcChrValEpn4	'참조문자값설명4

							If val = "DV_CO_CD" Then		'택배사코드
								rw cd & "||" & cdNm
							Else
								rw "그룹코드 : " & grpCd
								rw "언어코드 : " & langCd
								rw "코드 : " & cd
								rw "코드명 : " & cdNm
								rw "코드설명 : " & cdEpn
								rw "정렬순서 : " & sortSeq
								rw "참조문자값설명1 : " & refcChrValEpn1
								rw "참조문자값설명2 : " & refcChrValEpn2
								rw "참조문자값설명3 : " & refcChrValEpn3
								rw "참조문자값설명4 : " & refcChrValEpn4
								rw "--------------------------------------"
							End If
						Next

					Set datalist = nothing
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상세조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 등록
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
			iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & html2db(Err.Description)
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
							epdNo			= datalist.get(i).epdNo				'업체상품번호
							spdNo			= datalist.get(i).spdNo				'판매자상품번호
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCrlf
						strSql = strSql & " SET lotteonregdate = getdate()" & VbCrlf
						If (spdNo <> "") Then
'							strSql = strSql & "	, lotteonStatCd = '3'"& VbCRLF					'승인대기
							strSql = strSql & "	, lotteonStatCd = '7'"& VbCRLF					'등록완료
						Else
							strSql = strSql & "	, lotteonStatCd = '1'"& VbCRLF					'전송시도
						End If
						strSql = strSql & " ,lotteonGoodNo = '" & spdNo & "'" & VbCrlf
						strSql = strSql & " ,lotteonlastupdate = getdate()"
						strSql = strSql & " ,lotteonPrice = '"&imustprice&"' " & VbCrlf
						strSql = strSql & " ,lotteonsellYn = 'Y' "& VbCrlf
						strSql = strSql & " ,accFailCNT = 0" & VbCrlf                 ''실패회수 초기화
						strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
						strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_regitem R" & VbCrlf
						strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
						strSql = strSql & " where R.itemid = " & iitemid
						dbget.execute strSql
						iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
					Else
						iErrStr = "ERR||"&iitemid&"||"&resultMessage&"(Err1.상품등록)"
					End If
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.상품등록)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(상품등록)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-REG-002]"
		End If
	Set objXML= nothing
End Function

'상품 상세조회
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
			iErrStr = "ERR||"&iitemid&"||실패[상품조회] " & html2db(Err.Description)
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

					lotteonsellYn = strObj.data.slStatCd	'판매상태코드 [공통코드 : SL_STAT_CD] | END : 판매종료, SALE : 판매중, SOUT : 품절, STP : 판매중지
					'rw strObj.data.slStatRsnCd				'상품판매상태사유코드 [공통코드 : SL_STAT_RSN_CD] | SOUT_STK : 재고수량 0 품절처리 ,SOUT_LMT : 한정수량 0 품절처리 ,SOUT_ADMRSOUT_ADMR : 관리자 수동품절처리 ,SOUT_ITM : 모든단품품절로 상품 품절처리 ,SOUT_RSV : 예약상품 품절처리 ,SOUT_THDY : 명절 품절예약 ,STP_TNS : TNS(금칙어/신고) 패널티 판매중지 ,STP_CTL : 카탈로그에 의한 판매중지 ,STP_MNTR : 상품정보 모니터링 부적합 처리 ,STP_DV : 배송서비스 패널티 판매중지 ,STP_UNAPRV : 상품미승인 상태에 의한 판매중지처리 ,END_TR : 거래처계약종료로 인한 판매종료 처리 ,END_ADMR : 관리자 수동 판매종료 처리 ,END_TNS : TNS(위해상품) 판매종료 처리
					fnlAprvYn = strObj.data.pdAprvStatInfo.fnlAprvYn    '최종승인여부
					fnlAprvYnStr = Chkiif(fnlAprvYn="Y", "승인완료", "승인전")
					Set itmLst = strObj.data.itmLst		'단품목록
						For i=0 to itmLst.length-1
							outmallSellyn = ""
							outmallOptCode	= itmLst.get(i).sitmNo			'판매자단품번호
							outmallOptName	= itmLst.get(i).sitmNm			'판매자단품명
							If itmLst.get(i).slStatCd = "SALE" Then			'판매상태코드 [공통코드 : SL_STAT_CD] | SALE : 판매중, SOUT : 품절
								outmallSellyn = "Y"
							Else
								outmallSellyn = "N"
							End If
'							rw itmLst.get(i).slPrc			'판매가
							outmalllimitno	= itmLst.get(i).stkQty			'재고수량 | 재고관리여부가 Y인 경우에는 필수값

							strSql = " INSERT INTO db_item.dbo.tbl_OutMall_regedoption"
							strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outMallSellyn, outmalllimityn, outMallLimitNo)"
							strSql = strSql & " VALUES ("&iitemid
							If i = 0 AND outmallOptName = "단일상품" Then
								strSql = strSql & " ,'0000'"
							Else
								strSql = strSql & " ,'"& i &"'" ''임시로 롯데 코드 넣음 //2013/04/01
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
					iErrStr =  "OK||"&iitemid&"||성공("&fnlAprvYnStr&")"
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.상품조회)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(상품조회)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-CHKSTAT-002]"
		End If
	Set objXML= nothing
End Function

'승인 상품 수정
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
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
						Next
					Set datalist = nothing

					If resultCode = "0000" Then
						strSql = ""
						strSql = strSql & " UPDATE R"
						strSql = strSql & "	Set regitemname = '"&html2db(iitemname)&"'" & vbcrlf
						strSql = strSql & "	From db_etcmall.dbo.tbl_lotteon_regItem  R"
						strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||성공(상품수정)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& resultMessage &"실패(상품수정)"
					End If

				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.상품수정)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(상품수정)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-002]"
		End If
	Set objXML= nothing
End Function

'승인 상품 수정..호출은 위와 같으나 너무 느려서 하나로 통합해봄..2020-05-07
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
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
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
						iErrStr =  "OK||"&iitemid&"||성공(상품수정)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& resultMessage &"실패(상품수정)"
					End If

				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.상품수정)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(상품수정)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-EDIT-002]"
		End If
	Set objXML= nothing
End Function

'상품 재고 수정
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
			iErrStr = "ERR||"&iitemid&"||실패[재고] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
							sitmNo			= datalist.get(i).sitmNo			'판매자단품번호
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
						iErrStr =  "OK||"&iitemid&"||수정성공(재고)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"실패(재고)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-QTY-002]"
		End If
	Set objXML= nothing
End Function

'상품 가격 수정
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
			iErrStr = "ERR||"&iitemid&"||실패[가격] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
							sitmNo			= datalist.get(i).sitmNo			'판매자단품번호
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
						'// 상품가격정보 수정
						strSql = ""
						strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lotteon_regitem  " & VbCRLF
						strSql = strSql & "	SET lotteonLastUpdate=getdate() " & VbCRLF
						strSql = strSql & "	, lotteonPrice = " & imustprice & VbCRLF
						strSql = strSql & "	,accFailCnt = 0"& VbCRLF
						strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
						dbget.Execute(strSql)
						iErrStr =  "OK||"&iitemid&"||수정성공(상품가격)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"실패(상품가격)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-PRICE-002]"
		End If
	Set objXML= nothing
End Function

'상품 판매상태 변경
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
			iErrStr = "ERR||"&iitemid&"||실패[판매상태] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
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
							iErrStr =  "OK||"&iitemid&"||판매(상태변경)"
						ElseIf ichgSellyn = "N" Then
							strSql = ""
							strSql = strSql & " UPDATE R"
							strSql = strSql & "	Set lotteonSellYn = 'N'"
							strSql = strSql & "	,accFailCnt = 0"
							strSql = strSql & "	,lotteonLastUpdate = getdate()"
							strSql = strSql & "	From db_etcmall.dbo.tbl_lotteon_regitem  R"
							strSql = strSql & " WHERE R.itemid = '" & iitemid & "'"
							dbget.Execute(strSql)
							iErrStr =  "OK||"&iitemid&"||품절처리(상태변경)"
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
							iErrStr = "OK||"&iitemid&"||판매종료"
						End If
					Else
						If InStr(resultMessage, "중복상품으로 인한 판매중지가") Then
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
							strSql = strSql & "  	INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid, mallgubun, bigo) VALUES('"&iitemid&"','"&CMALLNAME&"', '존재하지 않거나 이미지 오류') "
							strSql = strSql & "  END "
							dbget.Execute strSql
							iErrStr = "OK||"&iitemid&"||판매중지(상태변경)/관리자 종료처리"
						Else
							iErrStr = "ERR||"&iitemid&"||"&resultMessage&"(Err1.판매상태)"
						End If
					End If
				Else
					If Err.number <> 0 Then
						iErrStr = "ERR||"&iitemid&"||"&Err.Description&"(ERR2.판매상태)"
					Else
						iErrStr = "ERR||"&iitemid&"||"& html2db(resultmsg) &"(판매상태)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-EDITSELLYN-002]"
		End If
	Set objXML= nothing
End Function

'단품 판매상태 변경
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
			iErrStr = "ERR||"&iitemid&"||실패[단품상태] " & html2db(Err.Description)
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
							resultCode		= datalist.get(i).resultCode		'처리결과
							resultMessage	= datalist.get(i).resultMessage		'처리메세지
							sitmNo			= datalist.get(i).sitmNo			'판매자단품번호
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
						iErrStr =  "OK||"&iitemid&"||수정성공(단품상태)"
					Else
						iErrStr =  "ERR||"&iitemid&"||"& sitmNo &"실패(단품상태)"
					End If
				End If
			Set strObj = nothing
		Else
			iErrStr = "ERR||"&iitemid&"||Lotteon 결과 분석 중에 오류가 발생했습니다.[ERR-OPTSTAT-002]"
		End If
	Set objXML= nothing
End Function

'내부 처리 완료
Function fnlotteonTenConfirmOrder(vOdNo, vodSeq, vProcSeq)
	Dim objXML, xmlDOM, iRbody, strSql
	Dim obj, strParam, strObj
	Dim vDvTrcStatDttm, returnCode
	vDvTrcStatDttm = FormatDate(now(), "00000000000000")

	Set obj = jsObject()
		Set obj("ifCompleteList")= jsArray()									'내부처리완료목록
			Set obj("ifCompleteList")(null) = jsObject()
				obj("ifCompleteList")(null)("dvRtrvDvsCd") = "DV"					'#배송회수구분코드 | DV:배송, RTRV:회수
				obj("ifCompleteList")(null)("odNo") = vOdNo							'#주문번호 : 주문테이블의 PK속성
				obj("ifCompleteList")(null)("odSeq") = vodSeq						'#주문순번 : 주문내역에 대해서 단품별로 부여되는 속성값 1
				obj("ifCompleteList")(null)("procSeq") = vProcSeq					'#처리순번 : Default 1 단품단위로 처리순서값을 정의함. 최초 입력시 1 이고 클레임이 발생할 경우 1씩 증가함
				obj("ifCompleteList")(null)("orglProcSeq") = ""						'o클레임인 경우 필수 원처리순번 : 클레임이 발생했을 경우 상위처리순번
				obj("ifCompleteList")(null)("clmNo") = ""							'o클레임인 경우 필수 클레임번호 : 클레임이 발생했을 경우의 번호
				obj("ifCompleteList")(null)("ifCplYN") = "Y"						'#연동완료여부(Y/N)
				obj("ifCompleteList")(null)("ifFlRsnCnts") = ""						'#연동실패 사유내역
		 		strParam = obj.jsString
	Set obj = nothing

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerIfCompleteInform", false				'출고/회수지시 연동완료 통보
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[내부확인] " & html2db(Err.Description)
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
						rw vOdNo & " 오류 : " & strObj.data.rsltMsg
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
		Set obj("deliveryProgressStateList")= jsArray()									'배송상태목록
			Set obj("deliveryProgressStateList")(null) = jsObject()
				obj("deliveryProgressStateList")(null)("odNo") = vOdNo						'#주문번호 : 주문테이블의 PK속성
				obj("deliveryProgressStateList")(null)("odSeq") = vodSeq					'#주문순번 : 주문내역에 대해서 단품별로 부여되는 속성값 1
				obj("deliveryProgressStateList")(null)("procSeq") = vProcSeq				'#처리순번 : Default 1 단품단위로 처리순서값을 정의함. 최초 입력시 1 이고 클레임이 발생할 경우 1씩 증가함
				obj("deliveryProgressStateList")(null)("odPrgsStepCd") = "12"				'#주문진행단계 | 11 : 출고지시, 12 : 상품준비, 13 : 발송완료, 14 : 배송완료, 15 : 수취완료, 23 : 회수지시, 24 : 회수진행, 25 : 회수완료, 26 : 회수확정
				obj("deliveryProgressStateList")(null)("dvTrcStatDttm") = vDvTrcStatDttm	'#배송상태발생일시
				obj("deliveryProgressStateList")(null)("spdNo") = vSpdNo					'#상품번호 : 롯데ON에서 관리되는 상품번호
				obj("deliveryProgressStateList")(null)("sitmNo") = vSitmNo					'#단품번호 : 롯데ON에서 관리되는 단품번호
				obj("deliveryProgressStateList")(null)("slQty") = vSlQty					'#수량 : 단품에 대한 주문수량
		 		strParam = obj.jsString
	Set obj = nothing

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", APIURL & "/v1/openapi/delivery/v1/SellerDeliveryProgressStateInform", false		'배송상태통보
		objXML.setRequestHeader "Authorization", "Bearer " & APIkey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[주문확인] " & html2db(Err.Description)
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
						rw vOdNo & " 오류 : " & strObj.data.rsltMsg
					End If
				End If
			Set strObj = nothing
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

'롯데ON배송상태 조회
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
		obj("srchStrtDt") = searchDate & "000000"			'#검색시작일자 yyyymmddhhmmss 배송지시생성일시
		obj("srchEndDt") =  searchDate & "235959"			'#검색종료일자 yyyymmddhhmmss
		obj("odNo") = vOrderNo			'주문번호
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
			iErrStr = "ERR||"&iitemid&"||실패[배송상태] " & html2db(Err.Description)
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
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>
