<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "homeplus"
CONST CUPJODLVVALID = TRUE		''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5			'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CHomeplusItem
	Public FItemid
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FisUsing
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FoptionCnt
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FHomeplusGoodNo
	Public FHomeplusprice		
	Public FHomeplusSellYn	
	Public FregedOptCnt       
	Public FaccFailCNT        
	Public FlastErrStr        
	Public Fdeliverytype      
	Public FrequireMakeDay    
	Public FinfoDiv       	
	Public Fsafetyyn      	
	Public FsafetyDiv     	
	Public FsafetyNum     	
	Public FmaySoldOut    	

	Public FHomeplusStatCD
	Public FhDIVISION
	Public FhGROUP
	Public FhDEPT
	Public FhCLASS
	Public FhSUBCLASS
	Public FdepthCode
	Public FbrandDepthCode

	Public MustPrice
	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Fregitemname

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO))
	End Function

	'// Homeplus �Ǹſ��� ��ȯ
	Public Function getHomeplusSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getHomeplusSellYn = "Y"
			Else
				getHomeplusSellYn = "N"
			End If
		Else
			getHomeplusSellYn = "N"
		End If
	End Function

	public function GetHomeplusLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5
		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetHomeplusLmtQty = 0
			Else
				GetHomeplusLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetHomeplusLmtQty = 999
		End If
	End Function

    Function getHomeplusSuplyPrice(optaddprice)
		getHomeplusSuplyPrice= cLng((MustPrice+optaddprice)*0.89)
    End Function

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim p, strRst, arrData, arrTmp
		If trim(Fkeywords) = "" Then Exit Function
		strRst = ""
		Fkeywords = replace(Fkeywords, ",,", ",")

		If instr(Fkeywords, ",") > 1 Then
			arrData = Split(Fkeywords, ",")
			arrTmp = FnDistinctData(arrData)
			strRst = "<TAGS>"
			For p=0 to Ubound(arrTmp)-1
				strRst = strRst & "<item><![CDATA["&arrTmp(p)&"]]></item>"
			Next
			strRst = strRst & "</TAGS>"
		End If
		getItemKeyword = strRst
	End Function

	'�迭���� �ߺ��� ����
	Function FnDistinctData(ByVal aData)
		Dim dicObj, items, returnValue
		Set dicObj = CreateObject("Scripting.dictionary")
			dicObj.removeall
			dicObj.CompareMode = 0
			'loop�� ���鼭 ���� �迭�� �ִ��� �˻� �� Add
			For Each items In aData
				If not dicObj.Exists(items) Then dicObj.Add items, items
			Next

			returnValue = dicObj.keys
		Set dicObj = Nothing
		FnDistinctData = returnValue
	End Function

	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getHomeplusOptionParamToReg
		Dim strSql, strRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		strRst = ""
		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
'rw strSql
'response.end
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''���ϻ�ǰ
					FItemOption = "0000"
					optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					itemSu = GetHomeplusLmtQty
				Else
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo

					If rsget("optnmLen")>100 then
					    optionname=DdotFormat(optionname,50)
					End If
				End If
				strRst = strRst &"<ITEM>"
				strRst = strRst &"	<s_ITEMNO>"&FItemOption&"</s_ITEMNO>"							'##*��ü �����۹�ȣ / ��ü�� �ش� ������(�ɼ�) ��ȣ ���߿� ProductResult������ ���ϱ� ���� �Է��Ͽ� �ش�.
				strRst = strRst &"	<i_SIZE>1</i_SIZE>"												'##*Size(Amos) / 1���� ���� 1,2,3,4����.)�ش� ������ ������ ��ü���� �����س����ñ� �ٶ��ϴ�. �ٸ� API���� ���˴ϴ�. I_ITEMNO+I_SIZE�� Ű ������ ��� �Ǿ� ���ϴ�.
				strRst = strRst &"	<s_OPTION_NAME><![CDATA["&optionname&"]]></s_OPTION_NAME>"		'##*�ɼǸ�
				strRst = strRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'������ / 1: WEB ���� 3: ���� �� ��(Default)�������� �� ��� 1�� ����
				strRst = strRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'������ / �������� 3���� ������ ��� ���� ���õȴ�
				strRst = strRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*�ǸŰ�
				strRst = strRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*���ް�(VAT����)
'				strRst = strRst &"	<i_ACCUMULATION_RATE></i_ACCUMULATION_RATE>"						'��ǰ�������� / ��ǰ�� FMC������
'				strRst = strRst &"	<d_RELEASE_DATE></d_RELEASE_DATE>"									'������� / ������� (YYYYMMDD)
				strRst = strRst &"</ITEM>"
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getHomeplusOptionParamToReg = strRst
	End Function

	'// ��ǰ����: �ɼ� �Ķ���� ����(��ǰ������)
	Public Function getHomeplusOptionParamToEDT
		Dim strSql, sRst, itemSu, itemoption, optionname, optaddprice
		Dim GetTenTenMargin, i, arrRows, sellstat
		Dim isOptionExists, notitemId, notmakerid
		Dim optiontypename, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, DelOpt

		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_homeplus 'homeplus'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close
		isOptionExists = isArray(arrRows)

		strSql = "SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'homeplus' and itemid =" & Fitemid
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notitemId = rsget("cnt")
		End If
		rsget.close

		strSql = "SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item as i join [db_temp].dbo.tbl_jaehyumall_not_in_makerid as m on i.makerid = m.makerid where i.itemid = "& Fitemid&" and m.mallgubun = 'homeplus'"
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			notmakerid = rsget("cnt")
		End If
		rsget.close

		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				itemoption			= ArrRows(1,i)
				optiontypename		= ArrRows(2,i)
'				 optionname			= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
				optionname			= Replace(db2Html(ArrRows(3,i)),":","")					'2015-05-15 11:16 ���� / ���տɼ��� �޸��� replace�ؼ� �����ƾ���..�� optionname �ּ���
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				preged				= ArrRows(11,i)
				optNameDiff			= ArrRows(12,i)
				forceExpired		= ArrRows(13,i)
				oopt				= ArrRows(14,i)
				ooptCd				= ArrRows(15,i)
				DelOpt				= ArrRows(16,i)
				optaddprice			= ArrRows(17,i)

				If IsSoldOut Then
					sellstat = 2
				Else
					If itemoption = "0000" AND UBound(ArrRows,2) = 0 Then
						optionname = oopt
						itemSu = GetHomeplusLmtQty
					Else
						If (optlimityn = "Y") Then
							If optLimit <= 5 Then
								itemSu = 0
							Else
								itemSu = optLimit - 5
							End If
						Else
							itemSu = 999
						End if
	
						If (DelOpt = 1) OR (isUsing = "N") OR (optsellyn = "N") OR (notitemId > 0) OR (notmakerid > 0) Then
							sellstat = 2
						Else
							sellstat = 1
						End If
					End If
					optionname = DdotFormat(optionname,50)
	
					GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
					If GetTenTenMargin < CMAXMARGIN Then
						MustPrice = Forgprice
					Else
						MustPrice = FSellCash
					End If
				End If

'rw itemoption
'rw ooptCd
'rw optionname
'rw itemSu
'rw MustPrice+optaddprice
'rw getHomeplusSuplyPrice(optaddprice)
'rw sellstat
'rw "------------"
				sRst = sRst &"<ITEM>"
				sRst = sRst &"	<s_ITEMNO>"&itemoption&"</s_ITEMNO>"							'*��ü �����۹�ȣ / ��ü�� �ش� ������(�ɼ�) ��ȣ ���߿� ProductResult������ ���ϱ� ���� �Է��Ͽ� �ش�.
				If preged = 1 Then
					sRst = sRst &"	<i_ITEMNO>"&ooptCd&"</i_ITEMNO>"							'�����۹�ȣ / �����Ǵ� �������̸� �ش� ���� �ݵ�� �Է��Ͽ� �ֽñ� �ٶ��ϴ� �ű� �߰��Ǵ� �������� ��쿡�� �Է����� ������
				End If
				sRst = sRst &"	<i_SIZE>1</i_SIZE>"												'*Size(Amos) / �ϴ��� ���� ����(AK ���� ������ ����Ʈ�� ����1���� ���� 1,2,3,4����.)�ش� ������ ������ ��ü���� �����س����ñ� �ٶ��ϴ�. �ٸ� API���� ���˴ϴ�.
				sRst = sRst &"	<s_OPTION_NAME><![CDATA["&optionname&"]]></s_OPTION_NAME>"		'*�ɼǸ�
				sRst = sRst &"	<i_STOCK_TYPE>1</i_STOCK_TYPE>"									'������ / 1: WEB ���� 3: ���� �� ��(Default)�������� �� ��� 1�� ����
				sRst = sRst &"	<i_LIBQTY>"&itemSu&"</i_LIBQTY>"								'������ / �������� 3���� ������ ��� ���� ���õȴ�
				sRst = sRst &"	<f_RETAILPRICE>"&MustPrice+optaddprice&"</f_RETAILPRICE>"		'*�ǸŰ� / ���ް� ������ ������ �Է��� ���� �ִ� ���ް��� ��� �ǸŰ��� ���� ���ް��� �ǸŰ��� �������ϴ�..API ���� ��ǰ�� ��� ���� �������� ���Ͽ��� �����Ƿ� �������� ���Ƿ� �������� ���ñ� �ٶ��ϴ�.
				sRst = sRst &"	<f_BUYPRICE>"&getHomeplusSuplyPrice(optaddprice)&"</f_BUYPRICE>"'*���ް�(VAT����)
				If preged = 1 Then
					sRst = sRst &"	<i_STATUS>"&sellstat&"</i_STATUS>"							'�Ǹ� ��/�Ǹ����� | 1: �Ǹ��� 2:�Ǹ�����, �ű� �߰��Ǵ� �������� �ڵ����� �Ǹ������� ó���˴ϴ�. �����Ǵ� �������� ��쿡�� �� �ʵ带 ����մϴ�.
				End If
'				sRst = sRst &"	<ACCUMULATION_RATE></ACCUMULATION_RATE>"						'��ǰ�������� / ��ǰ�� FMC������
'				sRst = sRst &"	<RELEASE_DATE></RELEASE_DATE>"									'������� / ������� (YYYYMMDD)
				sRst = sRst &"</ITEM>"
			Next
		End If
'response.end
		getHomeplusOptionParamToEDT = sRst
	End Function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getHomeplusAddImageParamToReg()
		Dim strRst, strSQL, i, strRst2
		strRst = ""
		strRst2 = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<s_IMG_BIG>"&FbasicImage&"</s_IMG_BIG>"		'*�⺻�̹��� URL | HTTP URL ����. �ش� �̹����� �ܺο��� �ٿ�ε� ������ URL�̾�� �Ѵ�(IP �� ��� ����, �������� ����)
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'�����̹��� URL | HTTP URL ����. ���� ���� ����� �� �ִ�. �ش� �̹����� �ܺο��� �ٿ�ε� ������ URL �̾�� �Ѵ�(IP�� ��� ����, �������� ����)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst2 = strRst2 &"	<item>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</item>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next

			If strRst2 <> "" Then
				strRst2 = "<s_IMG_SKCS1>"&strRst2&"</s_IMG_SKCS1>"
			End If
		End If
		rsget.Close
		getHomeplusAddImageParamToReg = strRst&strRst2
	End Function


	Function getItemNameFormat()
		Dim buf
		buf = "[�ٹ�����]"&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getHomeplusItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><a href=""http://direct.homeplus.co.kr/app.exhibition.category.Category.ghs?comm=usr.category.inf&ctg_id=133459"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_homeplus.jpg""></a></center></p><br>")
		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
		End Select
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getHomeplusItemContParamToReg = strRst
	End Function

	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// ���߿ɼ�Ȯ��
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close
			If chkMultiOpt Then
				'// ���߿ɼ� �϶�
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
				rsget.Close
			Else
				'// ���Ͽɼ��� ��
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//��� ��ȯ
		checkTenItemOptionValid = chkRst
	End Function

	'// ��ǰ����: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ������)
	Public Function getHomeplusAddImageParamToEDT()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"<BASIC>"&FbasicImage&"</BASIC>"		'*�⺻�̹��� URL | HTTP URL ����. �ش� �̹����� �ܺο��� �ٿ�ε� ������ URL�̾�� �Ѵ�(IP �� ��� ����, �������� ����)
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'�����̹��� URL | HTTP URL ����. ���� ���� ����� �� �ִ�. �ش� �̹����� �ܺο��� �ٿ�ε� ������ URL �̾�� �Ѵ�(IP�� ��� ����, �������� ����)
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"		<EXTRA>http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</EXTRA>"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getHomeplusAddImageParamToEDT = strRst
	End Function

	'// ��ǰ��� XML ����
	Public Function getHomeplusItemRegXML()
		Dim strRst
		'���� ���� �� �ݺ�����Ʈ �Ǽ�
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:createNewProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*��ü��ǰ�ڵ� | ��ü���� �����ϴ� �ش� ��ǰ�� ���� Unique�� �ĺ� �ڵ�(API ��ǰ ������ ���Ͽ� �����Ұ�)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*��ǰ��(Web) | �� �Ǹ� ��ǰ��
'		strRst = strRst & "				<s_PREFIX>[�ٹ�����]</s_PREFIX>"						'##�� ���� | ��ǰ�� �տ� �ٴ� ����
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'������
		strRst = strRst & "				<s_MAK_CORP><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></s_MAK_CORP>"	'##*������
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)&"</s_ORIGN>"		'##*������
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*����ī�װ� DIVISION | �ֻ��� �з��ڵ�
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*����ī�װ� GROUP | DIVISION ���� �з� �ڵ�
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*����ī�װ� DEPT | GROUP ���� �з� �ڵ�
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*����ī�װ� CLASS | DEPT ���� �з� �ڵ�
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*����ī�װ� SUBCLASS | CLASS ���� �з� �ڵ�
		strRst = strRst & "				<s_STORENO>"							'##*����ī�װ� | String[] | ���õ�� ī�װ� ���� ���� ����� �� �ִ�. ���� ��ǰ�� ���õ� ī�װ�.
		If (FbrandDepthCode <> "") AND (FbrandDepthCode <> "0") Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If (FdepthCode <> "") AND (FdepthCode <> "0") Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO><item>134079</item></s_BRANDNO>"	'##�귣��ī�װ� | String[] | �귣�� ī�װ� ���� ���� ����� �� �ִ�
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'����
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##��ǰ�������� | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*��ǰ�󼼼���
		strRst = strRst & getHomeplusAddImageParamToReg							'##*�̹�������
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*�ǸŽ����� | YYYY-MM-DD
		strRst = strRst & "				<i_TAXCODE>"&CHKIIF(FVatInclude="N","0","1")&"</i_TAXCODE>"		'##*�������� | 0: �����, 1:����
		strRst = strRst & "				<ITEMS>"&getHomeplusOptionParamToReg&"</ITEMS>"					'*ITEM(�ɼ�) | ITEM ����. ��ǰ�� �ɼ��׸��� ������ �� ���� �Է��Ͽ��� �Ѵ�.
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##���λ�ǰ���� | Y: ���λ�ǰ, N: ���λ�ǰ �ƴ�(Default)
		strRst = strRst & getItemKeyword										'##�˻� ���Ǿ� | ��ǰ�˻� �� ��ǰ�� �̿ܿ� �ش� ��ǰ�� �˻��ǵ��� �˻� ����� ����
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##���ݺ񱳻���Ʈ ���⿩�� | ���ݺ� ����Ʈ�� �ش� ��ǰ�� ����� �� ����..Y: ���ݺ񱳻���Ʈ ����, N: ���ݺ񱳻���Ʈ �� ����(default)
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'������ü�ڵ� | ��ü ���� �� ���� ��ü�ڵ� �ʼ� ���� �ƴϸ�, �� �Է� �� �⺻��� ������ü �ڵ�� �ڵ��Է� ������ü �ڵ� ��� �� ������ü �ڵ� ��ϵ�
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##��ǰ����������� �ʵ����� �������� | true�̸� ���� false�̸� ���� �� �� false�� ��� FIELDS �����͸� ��Ȯ�� �Է��Ͽ� ���� �Ͽ��� �Ѵ�
		strRst = strRst & getHomeplusItemInfoCdToReg							'##��ǰ����������� �ʵ����� | ��ǰ�������� ��ø� ���� �ʵ�����
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:createNewProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
'response.write strRst
'response.end
		getHomeplusItemRegXML = strRst
	End Function

	'// ��ǰ���� XML ����
	Public Function getHomeplusItemEditXML()
		Dim strRst
		'���� ���� �� �ݺ�����Ʈ �Ǽ�
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProduct xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<Product>"
		strRst = strRst & "				<i_STYLE>"&FHomeplusGoodno&"</i_STYLE>"				'*��Ÿ�Ϲ�ȣ | ��ǰ��� �� ���� �� ��ü��ǰ�ڵ������� ���� �Ǵ� Ȩ�÷��� ��ǰ(��Ÿ��)��ȣ
		strRst = strRst & "				<PRODUCT_CODE>"&FItemid&"</PRODUCT_CODE>"				'##*��ü��ǰ�ڵ� | ��ü���� �����ϴ� �ش� ��ǰ�� ���� Unique�� �ĺ� �ڵ�(API ��ǰ ������ ���Ͽ� �����Ұ�)
		strRst = strRst & "				<s_POS_NAME><![CDATA["&Trim(getItemNameFormat)&"]]></s_POS_NAME>"	'##*��ǰ��(Web) | �� �Ǹ� ��ǰ��
'		strRst = strRst & "				<s_PREFIX>[�ٹ�����]</s_PREFIX>"						'##�� ���� | ��ǰ�� �տ� �ٴ� ����
		strRst = strRst & "				<s_DESIGN></s_DESIGN>"									'������
		strRst = strRst & "				<s_MAK_CORP><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></s_MAK_CORP>"	'##*������
		strRst = strRst & "				<s_ORIGN>"&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)&"</s_ORIGN>"		'##*������
		strRst = strRst & "				<DIVISION>"&FhDIVISION&"</DIVISION>"	'##*����ī�װ� DIVISION | �ֻ��� �з��ڵ�
		strRst = strRst & "				<GROUP>"&FhGROUP&"</GROUP>"				'##*����ī�װ� GROUP | DIVISION ���� �з� �ڵ�
		strRst = strRst & "				<DEPT>"&FhDEPT&"</DEPT>"				'##*����ī�װ� DEPT | GROUP ���� �з� �ڵ�
		strRst = strRst & "				<CLASS>"&FhCLASS&"</CLASS>"				'##*����ī�װ� CLASS | DEPT ���� �з� �ڵ�
		strRst = strRst & "				<SUBCLASS>"&FhSUBCLASS&"</SUBCLASS>"	'##*����ī�װ� SUBCLASS | CLASS ���� �з� �ڵ�
		strRst = strRst & "				<s_STORENO>"							'##*����ī�װ� | String[] | ���õ�� ī�װ� ���� ���� ����� �� �ִ�. ���� ��ǰ�� ���õ� ī�װ�.
		If FbrandDepthCode <> "" Then
		strRst = strRst & "					<item>"&FbrandDepthCode&"</item>"
		End If
		If FdepthCode <> "" Then
		strRst = strRst & "					<item>"&FdepthCode&"</item>"
		End If
		strRst = strRst & "				</s_STORENO>"
		strRst = strRst & "				<s_BRANDNO><item>134079</item></s_BRANDNO>"	'##�귣��ī�װ� | String[] | �귣�� ī�װ� ���� ���� ����� �� �ִ�
		strRst = strRst & "				<s_STUFF></s_STUFF>"					'����
		strRst = strRst & "				<i_DES_KIND>1</i_DES_KIND>"				'##��ǰ�������� | 0:TEXT (Default) 1:HTML
		strRst = strRst & "				<s_DES><![CDATA["&getHomeplusItemContParamToReg&"]]></s_DES>"	'##*��ǰ�󼼼���
		strRst = strRst & getHomeplusAddImageParamToReg							'##*�̹�������
		strRst = strRst & "				<i_IMAGE_UPDATE>1</i_IMAGE_UPDATE>"		'0 : �̹��� ������Ʈ �ȵ� 1: �̹��� ���� �ʿ�
		strRst = strRst & "				<d_SDATE>"&DATE()&"</d_SDATE>"			'##*�ǸŽ����� | YYYY-MM-DD
		strRst = strRst & "				<c_HARMFUL_YN>N</c_HARMFUL_YN>"			'##���λ�ǰ���� | Y: ���λ�ǰ, N: ���λ�ǰ �ƴ�(Default)
		strRst = strRst & getItemKeyword										'##�˻� ���Ǿ� | ��ǰ�˻� �� ��ǰ�� �̿ܿ� �ش� ��ǰ�� �˻��ǵ��� �˻� ����� ����
		strRst = strRst & "				<c_COOP_SEND_YN>Y</c_COOP_SEND_YN>"		'##���ݺ񱳻���Ʈ ���⿩�� | ���ݺ� ����Ʈ�� �ش� ��ǰ�� ����� �� ����..Y: ���ݺ񱳻���Ʈ ����, N: ���ݺ񱳻���Ʈ �� ����(default)
		strRst = strRst & "				<s_BRAND></s_BRAND>"					'Ȩ�÷��� ���� �����Ͽ� �ִ� �귣�� �̸� ���� �־��ش�.
'		strRst = strRst & "				<DELIVERY_SEQ></DELIVERY_SEQ>"			'������ü�ڵ� | ��ü ���� �� ���� ��ü�ڵ� �ʼ� ���� �ƴϸ�, �� �Է� �� �⺻��� ������ü �ڵ�� �ڵ��Է� ������ü �ڵ� ��� �� ������ü �ڵ� ��ϵ�
		strRst = strRst & "				<FIELD_SKIP>false</FIELD_SKIP>"			'##��ǰ����������� �ʵ����� �������� | true�̸� ���� false�̸� ���� �� �� false�� ��� FIELDS �����͸� ��Ȯ�� �Է��Ͽ� ���� �Ͽ��� �Ѵ�
		strRst = strRst & getHomeplusItemInfoCdToReg							'##��ǰ����������� �ʵ����� | ��ǰ�������� ��ø� ���� �ʵ�����
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</m:updateProduct>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
'response.write strRst
'response.end
		getHomeplusItemEditXML = strRst
	End Function

	'// ��ǰ �̹��� ���� XML ����
	Public Function getHomeplusItemEditImgXML
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateImage xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"
		strRst = strRst & getHomeplusAddImageParamToEDT							'##*�̹�������
		strRst = strRst & "		</m:updateImage>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditImgXML = strRst
	End Function

	Public Function getHomeplusItemEditOPTXML
		Dim strRst
		'���� ���� �� �ݺ�����Ʈ �Ǽ�
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:updateProductItem xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<I_STYLENO>"&FHomeplusGoodno&"</I_STYLENO>"		'*��Ÿ�Ϲ�ȣ
		strRst = strRst & getHomeplusOptionParamToEDT								'*������ | �߰�/���� �� ������(�ɼ�)����.�߰� ������ ������ I_SIZE�� ���� ��ϵ� I_SIZE�� �޶�� �մϴ�.
		strRst = strRst & "		</m:updateProductItem>"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		getHomeplusItemEditOPTXML = strRst
	End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Public Function getHomeplusItemInfoCdToReg()
		Dim buf, strSQL, mallinfoCd, infoContent
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00000') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='N') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00007') AND (F.chkdiv ='Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00002') THEN '������������' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='99999') THEN '�Ƿ�' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00016') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '�����ŷ�����ȸ ���(�Һ��ں����ذ����)�� �ǰ��Ͽ� ������ �帳�ϴ�.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN I.itemname " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00003') AND (IC.safetyyn= 'Y') THEN '1' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00021') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN 'N' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00021') AND (IC.safetyyn= 'Y') THEN 'Y' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd <> '125018') THEN '�� ��ǰ�� KC ���������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00004') AND (IC.safetyyn= 'Y') AND (M.mallinfocd= '125018') THEN 'ȭ��ǰ���� ���� ��ǰ�Ǿ�ǰ����û �ɻ縦 ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00005') AND ((IC.safetyyn= 'N') OR IC.safetyyn= '') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00008') THEN '61502' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00011') THEN '61201' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00009') THEN '61301' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00014') THEN '61401' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00017') AND (F.chkdiv ='Y') THEN '�� ��ǰ�� ����������Ǹ� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00019') AND (F.chkdiv ='Y') THEN '��ǰ�������� ���� ���ԽŰ� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00020') AND (F.chkdiv ='Y') THEN '' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCdAdd='00018') AND (F.chkdiv ='Y') THEN infocontent  " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00006') THEN '0' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE convert(varchar(500),F.infocontent) END AS infocontent  " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'homeplus' and IC.itemid='"&FItemid&"'  " & vbcrlf
		strSQL = strSQL & " and not (F.chkdiv ='N' and (M.mallinfocd in ('134005', '133006', '130005', '113011', '101012', '102008', '107010', '108010', '103008', '104007', '105008', '106008', '135007', '131004', '131013', '131014', '132006', '115013', '115015', '115005', '116013', '111009'))) " & vbcrlf
		strSQL = strSQL & " and not (((IC.safetyyn= 'N') OR IC.safetyyn= '') and (M.mallinfocd in ('113016', '113017', '101003', '101004', '107015', '107016', '108017', '108018', '103003', '103004', '104003', '104004', '105003', '105004', '106003', '106004', '135003', '135004', '131010', '131011', '125018', '125019', '116017', '116018'))) " & vbcrlf
		rsget.Open strSQL,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<FIELDS>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    buf = buf &"	<item>"
				buf = buf & " 		<FILED_ID>"&mallinfoCd&"</FILED_ID>"
				buf = buf & " 		<VALUE><![CDATA["&infoContent&"]]></VALUE>"
				buf = buf &" 	</item>"
				rsget.MoveNext
			Loop
			buf = buf & "</FIELDS>"
		End If
		rsget.Close
		getHomeplusItemInfoCdToReg = buf
	End Function


	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CHomeplus
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMakerid
	Public FRectItemID

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	Public Sub getHomeplusNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' �ɼ� �߰��ݾ� �ִ°�� ��� �Ұ�. //�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            'addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"

            ''' 2013/05/29 Ư��ǰ�� ��� �Ұ� (ȭ��ǰ, ��ǰ��)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.homeplusStatCD,-9) as homeplusStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_brandCategory_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "												'�ö��/ȭ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_homeplus_regItem where homeplusStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHomeplusItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FHomeplusStatCD	= rsget("HomeplusStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FhDIVISION			= rsget("hDIVISION")
                FOneItem.FhGROUP			= rsget("hGROUP")
                FOneItem.FhDEPT				= rsget("hDEPT")
                FOneItem.FhCLASS			= rsget("hCLASS")
                FOneItem.FhSUBCLASS			= rsget("hSUBCLASS")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbrandDepthCode	= rsget("brandDepthCode")
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt<getdate()"
        addSql = addSql & "     and edDt>getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.HomeplusGoodNo, m.Homeplusprice, m.HomeplusSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.hDIVISION, '') as hDIVISION, isnull(pm.hGROUP, '') as hGROUP, isnull(pm.hDEPT, '') as hDEPT, isnull(pm.hCLASS, '') as hCLASS, isnull(pm.hSUBCLASS, '') as hSUBCLASS, isnull(pm.hCATEGORY_ID, '') as hCATEGORY_ID "
		strSql = strSql & "	, isnull(hm.depthCode, '') as depthCode, isnull(bm.depthCode, '') as brandDepthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','21','22') "
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_Homeplus_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as hm on hm.tenCateLarge=i.cate_large and hm.tenCateMid=i.cate_mid and hm.tenCateSmall=i.cate_small and c.infodiv = hm.infodiv "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.HomeplusGoodNo is Not Null "									'#��� ��ǰ��
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHomeplusItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FHomeplusGoodNo	= rsget("HomeplusGoodNo")
				FOneItem.FHomeplusprice		= rsget("Homeplusprice")
				FOneItem.FHomeplusSellYn	= rsget("HomeplusSellYn")
                FOneItem.FoptionCnt         = rsget("optionCnt")
                FOneItem.FregedOptCnt       = rsget("regedOptCnt")
                FOneItem.FaccFailCNT        = rsget("accFailCNT")
                FOneItem.FlastErrStr        = rsget("lastErrStr")
                FOneItem.Fdeliverytype      = rsget("deliverytype")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")
                FOneItem.FinfoDiv       	= rsget("infoDiv")
                FOneItem.Fsafetyyn      	= rsget("safetyyn")
                FOneItem.FsafetyDiv     	= rsget("safetyDiv")
                FOneItem.FsafetyNum     	= rsget("safetyNum")
                FOneItem.FmaySoldOut    	= rsget("maySoldOut")
                FOneItem.FhDIVISION			= rsget("hDIVISION")
                FOneItem.FhGROUP			= rsget("hGROUP")
                FOneItem.FhDEPT				= rsget("hDEPT")
                FOneItem.FhCLASS			= rsget("hCLASS")
                FOneItem.FhSUBCLASS			= rsget("hSUBCLASS")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbrandDepthCode	= rsget("brandDepthCode")
                FOneItem.Fregitemname		= rsget("regitemname")
		End If
		rsget.Close
	End Sub
End Class

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function getHomplusGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 homeplusgoodno FROM db_etcmall.dbo.tbl_homeplus_regitem WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
		getHomplusGoodNo = rsget("homeplusgoodno")
	rsget.Close
End Function
%>