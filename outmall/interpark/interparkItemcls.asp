<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "interpark"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST interparkAPIURL = "http://ipss1.interpark.com"
CONST CDEFALUT_STOCK = 999
CONST wapiURL = "http://wapi.10x10.co.kr"

Class CInterparkitem
	Public Fitemid
	Public Fitemname
	Public FMakerid
	Public Fbuycash
	Public Fsellcash
	Public Forgsellcash
	Public Fsourcearea
	Public Foptioncnt
	Public FRegdate
	Public Fsellyn
	Public Flimityn
	Public Flimitno
	Public Flimitsold
	Public Fcate_large
	Public Fcate_mid
	Public Fcate_small
	Public FMakerName
	Public FBrandName
	Public FBrandNameKor
	Public Fkeywords
	Public Fitemoption
	Public FItemOptionTypeName
	Public FItemOptionName
	Public Fbasicimage
	Public FregImageName
	Public Fmainimage
	Public Fmainimage2
	Public FInfoImage
	Public Fordercomment
	Public FItemContent
	Public Fvatinclude
	Public Finterparkdispcategory
	Public Fitemsize
	Public Fitemsource
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Foptaddprice
	Public FLastUpdate
	Public FSellEndDate
	Public FInfoImage1
	Public FInfoImage2
	Public FInfoImage3
	Public FInfoImage4
	Public FAddImage1
	Public FAddImage2
	Public FAddImage3
	Public FAddImage4
	Public FItemDiv
	Public Fisusing
	Public FInterparkPrdNo
	Public FmayiParkSellYn
	Public FdeliveryType
	Public FdefaultfreeBeasongLimit
	Public FSailYn
	Public FOrgPrice
	Public Finterparkregdate
	Public Fdeliverfixday
	Public Ffreight_min
	Public Ffreight_max
	Public FlastErrStr
	Public Fmayiparkprice
	Public FregOptCnt
	Public FMaySoldOut
	Public FbasicimageNm
	Public FAdultType

	Public FMayLimitSoldout
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
		End If
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//��� ��ȯ
		checkTenItemOptionValid = chkRst
	End Function

	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin, ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT isnull(outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner_addInfo "
		sqlStr = sqlStr & " WHERE partnerid = '"& CMALLNAME &"'  "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			outmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner "
		sqlStr = sqlStr & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : ����
		sqlStr = sqlStr & " and id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ownItemCnt = rsget("CNT")
		End If
		rsget.Close

		If specialPrice <> "" Then
			tmpPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			tmpPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < outmallstandardMargin Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	End Function

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	Function getiszeroWonSoldOut(iitemid, ilimityn)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
' o11st.FOneItem.FLimityn
		i = 0
		sqlStr = ""
		sqlStr = sqlStr & "SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option where itemid = '"&iitemid&"' and optaddprice > 0 "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			cnt = rsget("cnt")
		rsget.Close

		If cnt = 0 Then
			getiszeroWonSoldOut = "N"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemid, itemoption, optlimitno, optlimitsold "
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option  "
			sqlStr = sqlStr & " where itemid = '"&iitemid&"'  "
			sqlStr = sqlStr & " and optaddprice = 0 "
			sqlStr = sqlStr & " and isusing = 'Y' "
			sqlStr = sqlStr & " and optsellyn = 'Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				Do until rsget.EOF
					goptlimitno		= rsget("optlimitno")
					goptlimitsold	= rsget("optlimitsold")
					If (ilimityn = "Y") AND (goptlimitno - goptlimitsold > CMAXLIMITSELL) Then
						i = i + 1
					End If
					rsget.MoveNext
				Loop

				If (ilimityn = "Y") AND (i = 0) Then
					getiszeroWonSoldOut = "Y"
				ElseIf (ilimityn = "Y") AND (i > 0) Then
					getiszeroWonSoldOut = "N"
				Else
					getiszeroWonSoldOut = "N"
				End If
			Else
				getiszeroWonSoldOut = "Y"
			End If
			rsget.Close
		End If

		If getiszeroWonSoldOut = "Y" Then		'0�� ǰ���� ��ǰ���� ť�� �ױ�
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_outmall_API_Que "
			sqlStr = sqlStr & " (mallid, apiAction, itemid, priority, lastUserid) "
			sqlStr = sqlStr & " VALUES ('interpark', 'DELETE', '"&iitemid&"', 10, 'system') "
			dbget.Execute sqlStr
		End If
	End Function

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

	Public Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		'2017-07-03 ������ ��ǰ�� Ư�� ����
		buf = replace(buf,"��","")
		buf = replace(buf,"?","")
		buf = replace(buf,"��","")
		buf = replace(buf,"��","")
		buf = replace(buf,"��","")
		buf = replace(buf,"��","")
		buf = replace(buf,"��"," ")
		buf = replace(buf,"��","x")
		buf = replace(buf,"��",":")
		buf = replace(buf,"��","")
		buf = replace(buf,"��","'")
		buf = replace(buf,"`","")
		buf = replace(buf,"��",",")
		buf = replace(buf,"��","[")
		buf = replace(buf,"��","]")
		'2017-07-03 ������ ��ǰ�� Ư�� ���ų�
		buf = "[�ٹ�����] " & Replace(Replace(Replace(Replace(Replace(FBrandNameKor & " " & CStr(buf),"'",""),Chr(34),""),"<",""),">",""),"^","")
		getItemNameFormat = buf
	End Function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

    Public Function GetInterParkSaleStatTp
		If (IsSoldOut) Then
			if (FSellyn = "S") then
				GetInterParkSaleStatTp = "05"       ''ǰ��(02)     SellYN-S
			Else
				If (Fisusing = "N") Then
					GetInterParkSaleStatTp = "03"   ''�Ǹ�����
				Else
					GetInterParkSaleStatTp = "02"   ''"03"   ''�Ǹ�����(03) SellYN-N  //02�� ���� 2013/09/02
				End if
			End If
		ElseIf FMaySoldout = "Y" Then
			GetInterParkSaleStatTp = "02"
		Else
			GetInterParkSaleStatTp = "01"
		End If
    End Function

	Public Function GetInterParkLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5
		''Max 99999 -> 1000
		If (Flimityn = "Y") Then
			If (Flimitno-Flimitsold) < CLIMIT_SOLDOUT_NO then
				GetInterParkLmtQty = 0
			Else
				GetInterParkLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetInterParkLmtQty = 999
		End if
	End Function

    Public Function GetSellEndDateStr()
		GetSellEndDateStr = "99991231"
		If IsNULL(FSellEndDate) Then Exit Function
		FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    End Function

	Public Function IsTruckReturnDlvExists
		IsTruckReturnDlvExists = false
		If (FItemID = 240488) then
			IsTruckReturnDlvExists = false
			Exit Function
		End If

		If IsNULL(Ffreight_max) Then Exit Function
		If CStr(Ffreight_max = "") Then Exit Function

		IsTruckReturnDlvExists = (Fdeliverfixday="X") and (Ffreight_max>0)
	End Function

	Public Function getTruckReturnDlvPrice
		getTruckReturnDlvPrice = 0

		If (FItemID=240488) then
			getTruckReturnDlvPrice = 50000
			Exit Function
		End If

		getTruckReturnDlvPrice = CLNG(Ffreight_max*2)   '' ������ũ ����Ʈ�� ���� ��������.. ���ƾ� 2���?
	End Function

	Public Function getInterparkContParamToReg()
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
		strRst = strRst & "<p align='center'><a href='http://www.interpark.com/display/sellerAllProduct.do?_method=main&sc.entrNo=3000010614&sc.supplyCtrtSeq=2&mid1=middle&mid2=seller&mid3=001#N_E_B_50_1_~' target='_blank'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_iPark.jpg'></a></p><br>"
		If Fitemsize <> "" Then
			strRst = strRst & "- ������ : " & Fitemsize & "<br>"
		End if

		If Fitemsource <> "" Then
			strRst = strRst & "- ��� : " &  Fitemsource & "<br>"
		End If
		strRst = strRst & Replace(Replace(FItemContent,"",""),"","")

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=http://webimage.10x10.co.kr/image/main/" & GetImageSubFolderByItemid(FItemID) & "/" & Fmainimage & "><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=http://webimage.10x10.co.kr/image/main2/" & GetImageSubFolderByItemid(FItemID) & "/" & Fmainimage2 & "><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")
		getInterparkContParamToReg = strRst
	End Function

	Public Function GetInterParkentrPoint()
		GetInterParkentrPoint = CLng(Fsellcash*0.01)
		If (GetInterParkentrPoint < 10) Then GetInterParkentrPoint = 0
		If (Fsellcash < 1000) Then GetInterParkentrPoint = 0	'õ���̸��ǻ�ǰ�� ��������Ʈ ����� �Ұ��մϴ�.
		GetInterParkentrPoint = 0	'2013/02/07 ��������Ʈ����
	End Function

	Public Function getInterparkOptParamtoREG
		Dim sqlStr, optLimit, itemoption
		Dim optArrRows, optlp, optlpName, optlpCode, optlpSu, optlpUsing, optlpStr, buf, optstr
		Dim ioptNameBuf, ioptCodeBuf, ioptAddPrice, ioptLimitNo, ioptTypeName

		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
		sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
				If optLimit > 0 Then
			        ioptTypeName	= Replace(Replace(Trim(rsget("optionTypeName"))," ",""),"����","����")
					ioptCodeBuf		= ioptCodeBuf & rsget("itemoption") & ","
					ioptNameBuf		= ioptNameBuf & Replace(Replace(Replace(Replace(Trim(rsget("optionname")),",",".")," ",""),"<","("),">",")") & ","  ''�ɼǳ��뿡 ���� ������ �ȵ�.//������ �ɼ� �����Ϳ� ������
					ioptAddPrice	= ioptAddPrice & CStr(rsget("optaddprice")) & ","
					ioptLimitNo		= ioptLimitNo & CStr(optLimit) & ","
				End If
				rsget.MoveNext
			Loop
		Else
			getInterparkOptParamtoREG = ""
			rsget.Close
			Exit Function
		End If
		rsget.Close

		ioptNameBuf		= RightCommaDel(ioptNameBuf)
	    ioptCodeBuf		= RightCommaDel(ioptCodeBuf)
	    ioptAddPrice	= RightCommaDel(ioptAddPrice)
	    ioptLimitNo		= RightCommaDel(ioptLimitNo)

	    If (ioptTypeName="") then ioptTypeName="�ɼǸ�"
	    optstr = ioptTypeName & "<" & ioptNameBuf & ">"

        If (ioptLimitNo <> "") Then
            optstr = optstr & "����<" & ioptLimitNo & ">"
        End If
        optstr = optstr & "�߰��ݾ�<" & ioptAddPrice & ">"
        optstr = optstr & "�ɼ��ڵ�<" & ioptCodeBuf & ">"
		optstr = Replace(optstr, VbTab, "")

		If Fitemdiv = "06" Then
			buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
			buf = buf & "		<prdOption><![CDATA[{" & optstr & "}]]></prdOption>"
		Else
			buf = buf & "		<prdOption><![CDATA[" & optstr & "]]></prdOption>"
		End If
		getInterparkOptParamtoREG = buf
	End Function

	Public Function getInterparkOptParamtoEDT
		Dim sqlStr, optLimit, itemoption, limitNCnt, limitYCnt
		Dim optArrRows, optlp, optlpName, optlpCode, optlpSu, optlpUsing, optlpStr, buf, optstr
		Dim ioptNameBuf, ioptCodeBuf, ioptAddPrice, ioptLimitNo, ioptTypeName
		Dim notUbound
		Dim notCArr
		notUbound = ""
		limitYCnt = 0
		limitYCnt = 0

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 i.optioncnt, isnull(T.regedoptcnt, 0) as regedoptcnt "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_interpark_reg_Item as T on i.itemid = T.itemid "
		sqlStr = sqlStr & " WHERE i.itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			FOptionCnt = rsget("optioncnt")
			FregOptCnt = rsget("regedoptcnt")
		End If
		rsget.Close

		buf = ""
		If Fitemdiv = "06" Then
			buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
		End If
		If FOptionCnt = 0 AND FregOptCnt > 0 Then	'���� ��ǰ�ε�, ��ϴ�ÿ� �ɼ��� �־��� ���
			sqlStr = ""
		    sqlStr = sqlStr & " SELECT itemoption, outmallOptName "
		    sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption"
		    sqlStr = sqlStr & " WHERE itemid='"&FItemID&"' "
		    sqlStr = sqlStr & " and mallid = '"&CMALLNAME&"' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		    if not rsget.Eof then
		        optArrRows = rsget.getRows()
		    Else
		    	notUbound = "Y"
		    end if
		    rsget.close

			If notUbound = "" Then
			    For optlp =0 To UBound(optArrRows,2)
			    	optlpName	= optlpName & optArrRows(1,optlp) & ","
			    	optlpCode	= optlpCode & optArrRows(0,optlp) & ","
			    	optlpSu		= optlpSu & "0,"
			    	optlpUsing	= optlpUsing & "N,"
				Next
				optlpName	= RightCommaDel(optlpName)
				optlpCode	= RightCommaDel(optlpCode)
				optlpSu		= RightCommaDel(optlpSu)
				optlpUsing	= RightCommaDel(optlpUsing)

				optlpName	= "�ɼ�<" & optlpName & ">"
				optlpCode	= "�ɼ��ڵ�<" & optlpCode & ">"
				optlpSu		= "����<" & optlpSu & ">"
				optlpUsing	= "��뿩��<" & optlpUsing & ">"
				optlpStr = optlpName & optlpSu & optlpCode & optlpUsing
				optlpStr = Replace(optlpStr, VbTab, "")
				If Fitemdiv = "06" Then
					buf = buf & "		<prdOption><![CDATA[{" & optlpStr & "}]]></prdOption>"
				Else
					buf = buf & "		<prdOption><![CDATA[" & optlpStr & "]]></prdOption>"
				End If
				getInterparkOptParamtoEDT = buf
			Else
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_interpark_reg_item "
				sqlStr = sqlStr & " SET regedOptCnt = 0 "
				sqlStr = sqlStr & " WHERE itemid =" & FItemid
				dbget.Execute sqlStr
				getInterparkOptParamtoEDT = ""
			End If
		Else										'�� ��
			If FOptionCnt = 0 Then
				getInterparkOptParamtoEDT = ""
			ElseIf FItemid = "1422765" Then
				If FOptionCnt <> FregOptCnt Then
					sqlStr = ""
				    sqlStr = sqlStr & " SELECT itemoption, outmallOptName "
				    sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption"
				    sqlStr = sqlStr & " WHERE itemid='"&FItemID&"' "
				    sqlStr = sqlStr & " and mallid = '"&CMALLNAME&"' "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				    if not rsget.Eof then
				        optArrRows = rsget.getRows()
				    Else
				    	notUbound = "Y"
				    end if
				    rsget.close

					If notUbound = "" Then
					    For optlp =0 To UBound(optArrRows,2)
					    	optlpName	= optlpName & optArrRows(1,optlp) & ","
					    	optlpCode	= optlpCode & optArrRows(0,optlp) & ","
					    	optlpSu		= optlpSu & "0,"
					    	optlpUsing	= optlpUsing & "N,"
						Next
						optlpName	= RightCommaDel(optlpName)
						optlpCode	= RightCommaDel(optlpCode)
						optlpSu		= RightCommaDel(optlpSu)
						optlpUsing	= RightCommaDel(optlpUsing)

						optlpName	= "�ɼ�<" & optlpName & ">"
						optlpCode	= "�ɼ��ڵ�<" & optlpCode & ">"
						optlpSu		= "����<" & optlpSu & ">"
						optlpUsing	= "��뿩��<" & optlpUsing & ">"
						optlpStr = optlpName & optlpSu & optlpCode & optlpUsing
						optlpStr = Replace(optlpStr, VbTab, "")
'						If Fitemdiv = "06" Then
							buf = buf & "		<optPrirTp><![CDATA[01]]></optPrirTp>"
							buf = buf & "		<prdOption><![CDATA[{" & optlpStr & "}]]></prdOption>"
'						Else
'							buf = buf & "		<prdOption><![CDATA[" & optlpStr & "]]></prdOption>"
'						End If
						getInterparkOptParamtoEDT = buf
					Else
						sqlStr = ""
						sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_interpark_reg_item "
						sqlStr = sqlStr & " SET regedOptCnt = 0 "
						sqlStr = sqlStr & " WHERE itemid =" & FItemid
						dbget.Execute sqlStr
						getInterparkOptParamtoEDT = ""
					End If
				End If
			Else
				sqlStr = ""
				sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
				sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
				sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

				        ioptTypeName	= Replace(Replace(Trim(rsget("optionTypeName"))," ",""),"����","����")
						ioptCodeBuf		= ioptCodeBuf & rsget("itemoption") & ","
						ioptNameBuf		= ioptNameBuf & Replace(Replace(Replace(Replace(Trim(rsget("optionname")),",",".")," ",""),"<","("),">",")") & ","  ''�ɼǳ��뿡 ���� ������ �ȵ�.//������ �ɼ� �����Ϳ� ������
						ioptAddPrice	= ioptAddPrice & CStr(rsget("optaddprice")) & ","
						ioptLimitNo		= ioptLimitNo & CStr(optLimit) & ","
						If (optLimit = 0) Then
							optlpUsing	= optlpUsing & "N,"
							limitNCnt = limitNCnt + 1
						Else
							optlpUsing	= optlpUsing & "Y,"
							 limitYCnt =  limitYCnt + 1
						End If
						rsget.MoveNext
					Loop
				End If
				rsget.Close

				If FOptioncnt > 0 Then
					If limitYCnt = 0 Then
						FMayLimitSoldout = "Y"
					Else
						FMayLimitSoldout = "N"
					End If
				End If

				ioptNameBuf		= RightCommaDel(ioptNameBuf)
			    ioptCodeBuf		= RightCommaDel(ioptCodeBuf)
			    ioptAddPrice	= RightCommaDel(ioptAddPrice)
			    ioptLimitNo		= RightCommaDel(ioptLimitNo)
			    optlpUsing		= RightCommaDel(optlpUsing)

			    If (ioptTypeName="") then ioptTypeName="�ɼǸ�"
			    optstr = ioptTypeName & "<" & ioptNameBuf & ">"

                If (ioptLimitNo <> "") Then
                    optstr = optstr & "����<" & ioptLimitNo & ">"
                End If
                optstr = optstr & "�߰��ݾ�<" & ioptAddPrice & ">"
                optstr = optstr & "�ɼ��ڵ�<" & ioptCodeBuf & ">"
                optstr = optstr & "��뿩��<" & optlpUsing & ">"
				optstr = Replace(optstr, VbTab, "")
				If Fitemdiv = "06" Then
					buf = buf & "		<prdOption><![CDATA[{" & optstr & "}]]></prdOption>"
				Else
					buf = buf & "		<prdOption><![CDATA[" & optstr & "]]></prdOption>"
				End If
			End If
		End If
		getInterparkOptParamtoEDT = buf
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim keywordsBuf, keywordsStr, k
		keywordsBuf = Split(Fkeywords,",")
		For k = 0 to 2
			If UBound(keywordsBuf)> k Then keywordsStr = keywordsStr & Trim(keywordsBuf(k)) & ","
		Next
		keywordsStr = "�ٹ�����," & keywordsStr
		keywordsStr = RightCommaDel(keywordsStr)
		If (FItemID = 486220) or (FItemID = 486222) Then
			keywordsStr=""
		End If
		keywordsStr = Replace(keywordsStr,"'","")

		If stringCount(keywordsStr) > 100 Then
			keywordsStr = chrbyte("keywordsStr",100,"N")
		End If
		getItemKeyword = keywordsStr
	End Function

	Function stringCount(strString)
		Dim intPos, chrTemp, intLength
		'���ڿ� ���� �ʱ�ȭ
		intLength = 0
		intPos = 1

		'���ڿ� ���̸�ŭ ����
		while ( intPos <= Len( strString ) )
			'���ڿ��� �ѹ��ھ� ���Ѵ�
			chrTemp = ASC(Mid( strString, intPos, 1))
			if chrTemp < 0 then '������(-)�� ������ �ѱ���
				intLength = intLength + 2 '�ѱ��� ��� 2����Ʈ�� ���Ѵ�
			else
				intLength = intLength + 1 '�ѱ��� �ƴҰ�� 1����Ʈ�� ���Ѵ�
			end If
			intPos = intPos + 1
		wend
		stringCount = intLength
	End function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

	Public Function getBasicImage()
		If IsNULL(FbasicImageNm) or (FbasicImageNm="") Then Exit function
		getBasicImage = FbasicImageNm
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType = 2) or (FdeliveryType = 4) or (FdeliveryType = 5) Then
			IsFreeBeasong = True
		End If

		If (FSellcash >= 50000) Then IsFreeBeasong = True
	End Function

    Public Function getOrderCommentStr()
		Dim reStr
		reStr = ""
		If Not IsNULL(Fordercomment) Then
			If Len(Fordercomment) < 2 Then
				reStr = ""
			Else
				reStr = "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
			End If
		End If
		getOrderCommentStr = reStr
    End Function

    Public Function getInterparkAddImageParam()
    	Dim strRst, strSQL, i
    	strRst = ""
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&","
				End If
				rsget.MoveNext
				If i >= 4 Then Exit For
			Next
		End If
		rsget.Close
		getInterparkAddImageParam = RightCommaDel(strRst)
    End Function

	Public Function getInterparkItemsafetyReg
		Dim strSql, buf, safetyDiv, safetyNum, safetyYn, infoDiv, isElecCate, isLifeCate, isChildrenCate
		Dim certYN, bufLife, bufElec, bufChild, arrRows, notarrRows
		Dim newSafetyDiv, nLp, newDiv, newCertNo
		buf = ""
		bufLife = ""
		bufElec = ""
		bufChild = ""
		strSql = ""
		strSql = strSql & " SELECT TOP 1 " & vbcrlf
		strSql = strSql & " c.itemid, c.safetyYn, c.safetyDiv, isNULL(c.safetyNum, '') as safetyNum, c.infoDiv " & vbcrlf
		strSql = strSql & " , isnull(t.electric, '') as isElecCate, isnull(t.industrial, '') as isLifeCate, isnull(t.child, '') as isChildrenCate " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid " & vbcrlf
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_interpark_cate_mapping] m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall " & vbcrlf
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_interpark_category as t on m.CateKey = t.dispNo " & vbcrlf
		strSql = strSql & " WHERE i.itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			safetyYn		= rsget("safetyYn")
			safetyDiv		= rsget("safetyDiv")
			safetyNum		= rsget("safetyNum")
			isElecCate		= rsget("isElecCate")
			isLifeCate		= rsget("isLifeCate")
			isChildrenCate	= rsget("isChildrenCate")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
		strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows()
		Else
			notarrRows = "Y"
		End If
		rsget.Close

		If (isElecCate = "Y") OR (isLifeCate = "Y") OR (isChildrenCate = "Y") Then
			If notarrRows = "" Then		'���ȹ� ����� �����Ͷ�� ����� �ű�
				If safetyYn = "Y" Then
					For nLp =0 To UBound(arrRows,2)
				    	newDiv = ""
						Select Case arrRows(1,nLp)
							Case "10"		newDiv = "0201"		'�����ǰ > ��������
							Case "20"		newDiv = "0202"		'�����ǰ > ����Ȯ�� �Ű�
							Case "30"		newDiv = "0203"		'�����ǰ > ������ ���ռ� Ȯ��
							Case "40"		newDiv = "0101"		'��Ȱ��ǰ > ��������
							Case "50"		newDiv = "0102"		'��Ȱ��ǰ > ��������Ȯ��
							Case "60"		newDiv = "0104"		'��Ȱ��ǰ > ����ǰ��ǥ��
							Case "70"		newDiv = "0401"		'�����ǰ > ��������
							Case "80"		newDiv = "0402"		'�����ǰ > ����Ȯ��
							Case "90"		newDiv = "0403"		'�����ǰ > ������ ���ռ� Ȯ��
						End Select

						newCertNo = arrRows(0,nLp)
						If newCertNo = "x" Then
							newCertNo = ""
						End If

						If newDiv = "0201" OR newDiv = "0202" OR newDiv = "0203" Then
					    	bufElec = bufElec & "			<certInfo>"
					    	bufElec = bufElec & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufElec = bufElec & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufElec = bufElec & "			</certInfo>"
						ElseIf newDiv = "0101" OR newDiv = "0102" OR newDiv = "0104" Then
					    	bufLife = bufLife & "			<certInfo>"
					    	bufLife = bufLife & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufLife = bufLife & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufLife = bufLife & "			</certInfo>"
						ElseIf newDiv = "0401" OR newDiv = "0402" OR newDiv = "0403" Then
					    	bufChild = bufChild & "			<certInfo>"
					    	bufChild = bufChild & "				<certKind><![CDATA["&newDiv&"]]></certKind>"
					    	bufChild = bufChild & "				<certNo><![CDATA["&newCertNo&"]]></certNo>"
					    	bufChild = bufChild & "			</certInfo>"
						End If
					Next

					If (isElecCate = "Y") AND (isLifeCate = "Y") Then					'����� ��Ȱ�� �Ѵ� �ִ� ���� ī�װ�
						If bufElec <> "" OR bufLife <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufElec <> "" Then
								buf = buf & bufElec
							ElseIf bufLife <> "" Then
								buf = buf & bufLife
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isElecCate = "Y") AND (isChildrenCate = "Y") Then
						If bufElec <> "" OR bufChild <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufElec <> "" Then
								buf = buf & bufElec
							ElseIf bufChild <> "" Then
								buf = buf & bufChild
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isLifeCate = "Y") AND (isChildrenCate = "Y") Then
						If bufLife <> "" OR bufChild <> "" Then
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							If bufLife <> "" Then
								buf = buf & bufLife
							ElseIf bufChild <> "" Then
								buf = buf & bufChild
							End If
							buf = buf & "		</prdCertDetail>"
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isElecCate = "Y") Then
						If bufElec = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufElec
							buf = buf & "		</prdCertDetail>"
						End If
					ElseIf (isLifeCate = "Y") Then
						If bufLife = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufLife
							buf = buf & "		</prdCertDetail>"
						End If
					ElseIf (isChildrenCate = "Y") Then
						If bufChild = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						Else
					    	buf = buf & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
					    	buf = buf & "		<prdCertDetail>"
							buf = buf & bufChild
							buf = buf & "		</prdCertDetail>"
						End If
					End If
				Else
					buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
				End If
			Else						'���ȹ� ���� �� �� �����Ͷ�� ������ ������ �ű�
				If safetyNum = "" OR safetyYn = "N" then	'������ȣ�� ���ų� ������ �ƴ϶��
					certYN = "N"
				Else
					certYN = "Y"
				End If

				If certYN = "N" Then
					buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
				Else
				    If safetyDiv = "10" Then		'�츮�� ������������(KC��ũ)
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0101]]></certKind>"					'��Ȱ��ǰ] ��������Ȯ��
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
				    ElseIf safetyDiv = "20" Then	'�츮�� �����ǰ ��������
				    	bufElec = bufElec & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufElec = bufElec & "		<prdCertDetail>"
				    	bufElec = bufElec & "			<certInfo>"
				    	bufElec = bufElec & "				<certKind><![CDATA[0201]]></certKind>"					'[�����ǰ] ��������Ȯ��
				    	bufElec = bufElec & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufElec = bufElec & "			</certInfo>"
				    	bufElec = bufElec & "		</prdCertDetail>"
					ElseIf safetyDiv = "30" Then	'�츮�� KPS �������� ǥ��
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0104]]></certKind>"					'[��Ȱ��ǰ] ����ǰ��ǥ��
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					ElseIf safetyDiv = "40" Then	'�츮�� KPS �������� Ȯ�� ǥ��
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0102]]></certKind>"					'[��Ȱ��ǰ] ��������Ȯ��
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					ElseIf safetyDiv = "50" Then	'�츮�� KPS ��� ��ȣ���� ǥ��
				    	bufLife = bufLife & "		<prdCertStatus><![CDATA[Y]]></prdCertStatus>"
				    	bufLife = bufLife & "		<prdCertDetail>"
				    	bufLife = bufLife & "			<certInfo>"
				    	bufLife = bufLife & "				<certKind><![CDATA[0103]]></certKind>"					'[��Ȱ��ǰ] ��̺�ȣ����
				    	bufLife = bufLife & "				<certNo><![CDATA["&safetyNum&"]]></certNo>"
				    	bufLife = bufLife & "			</certInfo>"
				    	bufLife = bufLife & "		</prdCertDetail>"
					End If

					If (isElecCate = "Y") AND (isLifeCate = "Y") Then					'����� ��Ȱ�� �Ѵ� �ִ� ���� ī�װ�
						If bufElec <> "" Then
							buf = buf & bufElec
						ElseIf bufLife <> "" Then
							buf = buf & bufLife
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isElecCate = "Y") AND (isChildrenCate = "Y") Then
						If bufElec <> "" Then
							buf = buf & bufElec
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isLifeCate = "Y") AND (isChildrenCate = "Y") Then
						If bufLife <> "" Then
							buf = buf & bufLife
						Else
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						End If
					ElseIf (isElecCate = "Y") Then
						If bufElec = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						Else
							buf = buf & bufElec
						End If
					ElseIf (isLifeCate = "Y") Then
						If bufLife = "" Then
							buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
						Else
							buf = buf & bufLife
						End If
					ElseIf (isChildrenCate = "Y") Then
						buf = buf & "		<prdCertStatus><![CDATA[S]]></prdCertStatus>"			'��ǰ ���� �� ǥ��
					End If
				End If
			End If
		Else
			buf = buf & "		<prdCertStatus><![CDATA[N]]></prdCertStatus>"
		End If
		getInterparkItemsafetyReg = buf
'rw buf
	End Function

	'���� ��ǰǰ����� �ڵ� ���� 2012-11-12 ����
    Public Function getInterparkItemInfoCdToReg()
		Dim strSql, buf
		Dim mallinfoCd,infoContent,infotype
		'''IC.safetyyn => isNULL(IC.safetyyn,'N')
		'2014-05-15������ 00002 �߰�
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Interpark_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & "<prdinfoNoti>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infotype	= rsget("infotype")
			    infoContent = rsget("infoContent")

			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replace(infoContent, chr(31), "")
				End If

				buf = buf & "<info>"
				buf = buf & "	<infoSubNo><![CDATA["&mallinfoCd&"]]></infoSubNo>"
				buf = buf & "	<infoCd>"&infotype&"</infoCd>"
				buf = buf & "	<infoTx><![CDATA["&infoContent&"]]></infoTx>"
				buf = buf & "</info>"
				rsget.MoveNext
			Loop
			buf = buf & "</prdinfoNoti>"
		End If
		rsget.Close
		getInterparkItemInfoCdToReg = buf
    End Function

	Public Function getInterparkItemRegParameter()
		Dim strRst
	    strRst = ""
	    strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
	    strRst = strRst & "<result>"
	    strRst = strRst & "	<title>Interpark Product API</title>"
	    strRst = strRst & "	<description>��ǰ ���</description>"
	    strRst = strRst & "	<item>"
		strRst = strRst & "		<prdStat>01</prdStat>"																'��ǰ���� - ����ǰ:01, �߰��ǰ:02, ��ǰ��ǰ:03
		strRst = strRst & "		<shopNo>0000100000</shopNo>"														'������ũ ������ȣ (default - 0000100000)  | ������ȣ API��ü ����
		strRst = strRst & "		<omDispNo>" & Trim(Finterparkdispcategory) & "</omDispNo>" 							'������ũ �����ڵ�
	    strRst = strRst & "		<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"									'��ǰ�� - �ѱ� 60�� (����/���� 120��)
	    strRst = strRst & "		<hdelvMafcEntrNm><![CDATA["&CStr(FMakerName)&"]]></hdelvMafcEntrNm>"				'������ü��
	    strRst = strRst & "		<prdOriginTp><![CDATA["&GetSourcearea&"]]></prdOriginTp>"							'������
	    strRst = strRst & "		<taxTp>"&Chkiif(Fvatinclude="Y", "01", "02")&"</taxTp>"								'�ΰ��鼼��ǰ - ������ǰ:01, �鼼��ǰ:02, ������ǰ:03
	    strRst = strRst & "		<ordAgeRstrYn>"&Chkiif(IsAdultItem() = "Y", "Y", "N")&"</ordAgeRstrYn>"				'���ο�ǰ���� - ���ο�ǰ:Y, �Ϲݿ�ǰ:N
		strRst = strRst & "		<saleStatTp>01</saleStatTp>"														'�Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05, �����Ǹ�:09, ��ǰ����:98
	    strRst = strRst & "		<saleUnitcost>"&MustPrice&"</saleUnitcost>"											'�ǸŰ�
		strRst = strRst & "		<saleLmtQty>"&GetInterParkLmtQty&"</saleLmtQty>"									'�Ǹż��� - 99999 �� ���Ϸ� �Է�
		strRst = strRst & "		<saleStrDts>"&Replace(Left(CStr(now()),10),"-","")&"</saleStrDts>"					'�ǸŽ����� - yyyyMMdd => ȣ���� ��¥
		strRst = strRst & "		<saleEndDts>"&GetSellEndDateStr&"</saleEndDts>"										'�Ǹ������� - yyyyMMdd => 99991231 (������)
'2017-11-27 ������ ����..�� ��ǰ N����
'		strRst = strRst & "		<proddelvCostUseYn>"&Chkiif(Fdeliverytype="4", "Y", "N")&"</proddelvCostUseYn>"		'��ǰ��ۺ��뿩�� - ��ǰ��ۺ���:Y, ��ü��ۺ���å���:N
		strRst = strRst & "		<proddelvCostUseYn>N</proddelvCostUseYn>"											'��ǰ��ۺ��뿩�� - ��ǰ��ۺ���:Y, ��ü��ۺ���å���:N
	If IsTruckReturnDlvExists Then
		strRst = strRst & "		<prdrtnCostUseYn>Y</prdrtnCostUseYn>"												'��ǰ ��ǰ�ù�� ��뿩�� - ��ǰ��ǰ�ù����:Y, ��ü��ǰ�ù����:N
		strRst = strRst & "		<rtndelvCost>"&getTruckReturnDlvPrice&"</rtndelvCost>"								'��ǰ ��ǰ�ù��. prdrtnCostUseYn �� 'Y' �� ��� �ʼ���
	End If
		strRst = strRst & "		<prdBasisExplanEd><![CDATA["&getInterparkContParamToReg&"]]></prdBasisExplanEd>"	'��ǰ����
		strRst = strRst & "		<zoomImg><![CDATA["&Fbasicimage&"]]></zoomImg>"										'��ǥ�̹��� - ��ǥ�̹��� URL, ����/���� ����, JPG�� GIF�� ����
		strRst = strRst & "		<prdKeywd><![CDATA["&getItemKeyword&"]]></prdKeywd>"								'�����±� - �ִ� 4������, �޸��� ����
		strRst = strRst & "		<brandNm><![CDATA["&Fbrandname&"]]></brandNm>"										'�귣���
		strRst = strRst & "		<entrPoint>"&GetInterParkentrPoint&"</entrPoint>"									'��üPOINT - ��ü�ο� ����Ʈ �ݾ� �Է�, �ǸŰ��� �ִ� 10%���� ����
		strRst = strRst & "		<perordRstrQty>"& getOrderMaxNum &"</perordRstrQty>"								'1ȸ�� �ֹ� ���� ����
		strRst = strRst & "		<minOrdQty>1</minOrdQty>"															'�ּұ��ż��� - 1�� �̻� �Է�
		strRst = strRst & getInterparkOptParamtoREG
	If (Fitemdiv = "06") Then
		strRst = strRst & "		<inOpt>�ֹ����۹���</inOpt>"														'�Է��� �ɼ�. ex) ����ǰ�� �Է��ϼ���.
	End If
'2017-11-27 ������ ����..�� ��ǰ N���� �����߱⿡ �ϴ� �ʵ� �ּ�
'	If (Fdeliverytype = "4") Then
'		strRst = strRst & "		<delvCost>0</delvCost>"																'��ۺ� -��ǰ ��ۺ� �����϶� �ʼ�, 0�̸� ������
'	End If
		strRst = strRst & "		<delvAmtPayTpCom>"&Chkiif(FdeliveryType = "7", "01", "02")&"</delvAmtPayTpCom>"		'��ۺ� ���� ��� - ����:01, ����:02, ������ȯ���Ұ���:03 ��ǰ��ۺ� ����� ��� �ʼ�, �������϶�:02
    	strRst = strRst & "		<delvCostApplyTp>02</delvCostApplyTp>"												'��ۺ� ���� ��� - ����:01, ������:02
	If (IsFreeBeasong) Then
		strRst = strRst & "		<freedelvStdCnt>1</freedelvStdCnt>"													'�����۱��� ���� - ���ؼ��� �Է� ������� ���� ��� 0
	End If
		strRst = strRst & "		<jejuetcDelvCostUseYn>Y</jejuetcDelvCostUseYn>"										'���ֵ����갣��ۺ��뿩�� - Y : ���/����, N : ������
		strRst = strRst & "		<jejuDelvCost>3000</jejuDelvCost>"													'���ֹ�ۺ� - jejuetcDelvCostUseYn�� Y�϶� ���ֹ�ۺ�� �����갣�� �� �� �ϳ��� �ʼ�, 0�̸� ���ֹ�ۺ� 0��, null�̸� ������
		strRst = strRst & "		<etcDelvCost>3000</etcDelvCost>"													'�����갣��ۺ� - jejuetcDelvCostUseYn�� Y�϶� ���ֹ�ۺ�� �����갣�� �� �� �ϳ��� �ʼ�, 0�̸� �����갣��ۺ� 0��, null�̸� ������
		strRst = strRst & "		<spcaseEd><![CDATA[" & getOrderCommentStr & "]]></spcaseEd>"						'Ư�̻���
		strRst = strRst & "		<pointmUseYn>N</pointmUseYn>"														'����Ʈ����Ͽ��� - ����Ʈ����ǰ:Y, �Ϲݻ�ǰ:N �� 500�� �̸� ��ǰ�� ����� �Ұ����մϴ�. || 2013/02/07 ���� ���ſ��� ��������Ʈ ������ �پ ��а� ������ϱ�� �߰ŵ��~ �� ����� ������ ����ϱ�� �ؼ� ���ú��� ��������Ʈ�� �� ���ֽø� �� �� �����ϴ�~   ���� ���� �����̰� ���� �ʿ����� ���� ��û ���ٰ� �մϴ�~
		strRst = strRst & "		<ippSubmitYn>Y</ippSubmitYn>"														'���ݺ񱳵�Ͽ���
		strRst = strRst & "		<originPrdNo>"&CStr(FItemID)&"</originPrdNo>"										'��ǰ��ȣ
		strRst = strRst & "		<asInfo>������������</asInfo>"													'A/S���� | 2016-08-11 ������ �߰�..�ű԰����ʵ�
		strRst = strRst & "		<detailImg>"&getInterparkAddImageParam&"</detailImg>"								'���̹��� - ���̹��� URL, ����/���� ����, JPG�� GIF�� ���� �ִ� 4���� �̹�������, �޸�(,)�� �����Ͽ� ���.
		strRst = strRst & getInterparkItemsafetyReg()
 		strRst = strRst & getInterparkItemInfoCdToReg()
	    strRst = strRst & "	</item>"
	    strRst = strRst & "</result>"
		getInterparkItemRegParameter = strRst
	End Function

	Public Function getInterparkItemEditParameter()
	    Dim strRst
	    strRst = ""
	    strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
	    strRst = strRst & "<result>"
	    strRst = strRst & "	<title>Interpark Product API</title>"
	    strRst = strRst & "	<description>��ǰ ����</description>"
	    strRst = strRst & "	<item>"
	    strRst = strRst & "		<prdNo>"&FInterparkPrdNo&"</prdNo>"													'������ũ ��ǰ��ȣ
	    strRst = strRst & "		<prdStat>01</prdStat>"																'��ǰ���� - ����ǰ:01, �߰��ǰ:02, ��ǰ��ǰ:03
	    strRst = strRst & "		<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"									'��ǰ��
	    strRst = strRst & "		<hdelvMafcEntrNm><![CDATA["&CStr(FMakerName)&"]]></hdelvMafcEntrNm>"				'������ü��
	    strRst = strRst & "		<prdOriginTp><![CDATA["&GetSourcearea&"]]></prdOriginTp>"							'������
	    strRst = strRst & "		<taxTp>"&Chkiif(Fvatinclude="Y", "01", "02")&"</taxTp>"								'�ΰ��鼼��ǰ - ������ǰ:01, �鼼��ǰ:02, ������ǰ:03
	    strRst = strRst & "		<ordAgeRstrYn>"&Chkiif(IsAdultItem() = "Y", "Y", "N")&"</ordAgeRstrYn>"				'���ο�ǰ���� - ���ο�ǰ:Y, �Ϲݿ�ǰ:N
	    strRst = strRst & "		<saleStatTp>"&GetInterParkSaleStatTp&"</saleStatTp>"								'�Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05, �����Ǹ�:09, ��ǰ����:98
	    strRst = strRst & "		<saleUnitcost>"&MustPrice&"</saleUnitcost>"											'�ǸŰ�
		strRst = strRst & "		<saleLmtQty>"&GetInterParkLmtQty&"</saleLmtQty>"									'�Ǹż��� - 99999 �� ���Ϸ� �Է�
		'2018-04-17 ������..�����ÿ� �Ʒ� �ʵ� �ּ�
		'2018-11-01 ������..������ saleStrDts �ʵ���ٰ� ���� �߻����� �ּ� ����
		strRst = strRst & "		<saleStrDts>"&Replace(Left(CStr(now()),10),"-","")&"</saleStrDts>"						'�ǸŽ����� - yyyyMMdd => ȣ���� ��¥
		strRst = strRst & "		<saleEndDts>"&GetSellEndDateStr&"</saleEndDts>"										'�Ǹ������� - yyyyMMdd => 99991231 (������)
'2017-11-27 ������ ����..�� ��ǰ N����
'		strRst = strRst & "		<proddelvCostUseYn>"&Chkiif(Fdeliverytype="4", "Y", "N")&"</proddelvCostUseYn>"		'��ǰ��ۺ��뿩�� - ��ǰ��ۺ���:Y, ��ü��ۺ���å���:N
		strRst = strRst & "		<proddelvCostUseYn>N</proddelvCostUseYn>"											'��ǰ��ۺ��뿩�� - ��ǰ��ۺ���:Y, ��ü��ۺ���å���:N
	If IsTruckReturnDlvExists Then
		strRst = strRst & "		<prdrtnCostUseYn>Y</prdrtnCostUseYn>"												'��ǰ ��ǰ�ù�� ��뿩�� - ��ǰ��ǰ�ù����:Y, ��ü��ǰ�ù����:N
		strRst = strRst & "		<rtndelvCost>"&getTruckReturnDlvPrice&"</rtndelvCost>"								'��ǰ ��ǰ�ù��. prdrtnCostUseYn �� 'Y' �� ��� �ʼ���
	End If
		strRst = strRst & "		<prdBasisExplanEd><![CDATA["&getInterparkContParamToReg&"]]></prdBasisExplanEd>"	'��ǰ����
		strRst = strRst & "		<zoomImg><![CDATA["&Fbasicimage&"]]></zoomImg>"										'��ǥ�̹��� - ��ǥ�̹��� URL, ����/���� ����, JPG�� GIF�� ����
		strRst = strRst & "		<prdKeywd><![CDATA["&getItemKeyword&"]]></prdKeywd>"								'�����±� - �ִ� 4������, �޸��� ����
		strRst = strRst & "		<brandNm><![CDATA["&Fbrandname&"]]></brandNm>"										'�귣���
		strRst = strRst & "		<entrPoint>"&GetInterParkentrPoint&"</entrPoint>"									'��üPOINT - ��ü�ο� ����Ʈ �ݾ� �Է�, �ǸŰ��� �ִ� 10%���� ����
		strRst = strRst & "		<perordRstrQty>"& getOrderMaxNum &"</perordRstrQty>"								'1ȸ�� �ֹ� ���� ����
		strRst = strRst & "		<minOrdQty>1</minOrdQty>"															'�ּұ��ż��� - 1�� �̻� �Է�
		strRst = strRst & getInterparkOptParamtoEDT
	If (Fitemdiv = "06") Then
		strRst = strRst & "		<inOpt>�ֹ����۹���</inOpt>"														'�Է��� �ɼ�. ex) ����ǰ�� �Է��ϼ���.
	End If
'2017-11-27 ������ ����..�� ��ǰ N���� �����߱⿡ �ϴ� �ʵ� �ּ�
'	If (Fdeliverytype = "4") Then
'		strRst = strRst & "		<delvCost>0</delvCost>"																'��ۺ� -��ǰ ��ۺ� �����϶� �ʼ�, 0�̸� ������
'	End If
		strRst = strRst & "		<delvAmtPayTpCom>"&Chkiif(FdeliveryType = "7", "01", "02")&"</delvAmtPayTpCom>"		'��ۺ� ���� ��� - ����:01, ����:02, ������ȯ���Ұ���:03 ��ǰ��ۺ� ����� ��� �ʼ�, �������϶�:02
    	strRst = strRst & "		<delvCostApplyTp>02</delvCostApplyTp>"												'��ۺ� ���� ��� - ����:01, ������:02
	If (IsFreeBeasong) Then
		strRst = strRst & "		<freedelvStdCnt>1</freedelvStdCnt>"													'�����۱��� ���� - ���ؼ��� �Է� ������� ���� ��� 0
	End If
		strRst = strRst & "		<jejuetcDelvCostUseYn>Y</jejuetcDelvCostUseYn>"										'���ֵ����갣��ۺ��뿩�� - Y : ���/����, N : ������
		strRst = strRst & "		<jejuDelvCost>3000</jejuDelvCost>"													'���ֹ�ۺ� - jejuetcDelvCostUseYn�� Y�϶� ���ֹ�ۺ�� �����갣�� �� �� �ϳ��� �ʼ�, 0�̸� ���ֹ�ۺ� 0��, null�̸� ������
		strRst = strRst & "		<etcDelvCost>3000</etcDelvCost>"													'�����갣��ۺ� - jejuetcDelvCostUseYn�� Y�϶� ���ֹ�ۺ�� �����갣�� �� �� �ϳ��� �ʼ�, 0�̸� �����갣��ۺ� 0��, null�̸� ������
		strRst = strRst & "		<spcaseEd><![CDATA[" & getOrderCommentStr & "]]></spcaseEd>"						'Ư�̻���
		strRst = strRst & "		<pointmUseYn>N</pointmUseYn>"														'����Ʈ����Ͽ��� - ����Ʈ����ǰ:Y, �Ϲݻ�ǰ:N �� 500�� �̸� ��ǰ�� ����� �Ұ����մϴ�. || 2013/02/07 ���� ���ſ��� ��������Ʈ ������ �پ ��а� ������ϱ�� �߰ŵ��~ �� ����� ������ ����ϱ�� �ؼ� ���ú��� ��������Ʈ�� �� ���ֽø� �� �� �����ϴ�~   ���� ���� �����̰� ���� �ʿ����� ���� ��û ���ٰ� �մϴ�~
		strRst = strRst & "		<ippSubmitYn>Y</ippSubmitYn>"														'���ݺ񱳵�Ͽ���
		strRst = strRst & "		<originPrdNo>"&CStr(FItemID)&"</originPrdNo>"										'��ǰ��ȣ
		strRst = strRst & "		<asInfo>������������</asInfo>"													'A/S���� | 2016-08-11 ������ �߰�..�ű԰����ʵ�
	If isImageChanged Then
		strRst = strRst & "		<detailImg>"&getInterparkAddImageParam&"</detailImg>"								'���̹��� - ���̹��� URL, ����/���� ����, JPG�� GIF�� ���� �ִ� 4���� �̹�������, �޸�(,)�� �����Ͽ� ���.
		strRst = strRst & "		<imgUpdateYn>Y</imgUpdateYn>"														'�̹����������� - ��ǥ�̹���,���̹����� �������θ� ���� �մϴ�.Y : �̹��� ���� �ʿ�N : �̹��� ���� ���ʿ�(�⺻�� : N)��ǥ�̹����� ���̹��� �߿� �ϳ����̶� ������ �ʿ��� ��� Y�� �����ؾ� �մϴ�.
	End If
		strRst = strRst & getInterparkItemsafetyReg()
 		strRst = strRst & getInterparkItemInfoCdToReg()
	    strRst = strRst & "	</item>"
	    strRst = strRst & "</result>"
		getInterparkItemEditParameter = strRst
	End Function
End Class

Class CInterpark
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount
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

	Public Sub getInterparkNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,c.usinghtml, m.CateKey "
        strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
        strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName"
		strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
	    strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE s.itemid = i.itemid"
		strSql = strSql & " and s.itemid in ("
		strSql = strSql & " 	SELECT TOP " & CStr(FPageSize * FCurrPage) & " s.itemid "
		strSql = strSql & " 	FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i, [db_etcmall].[dbo].tbl_interpark_cate_mapping p "
		strSql = strSql & " 	WHERE s.itemid = i.itemid"
		strSql = strSql & "		and s.interparkregdate is NULL"
		strSql = strSql & "		and i.basicimage is not null"
		strSql = strSql & "		and i.cate_large <> '' "
		strSql = strSql & "		and i.cate_large <> '999' "
		strSql = strSql & "		and i.sellcash > 0"
		strSql = strSql & " 	and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'�ù�(�Ϲ�)
		strSql = strSql & "		and i.itemdiv in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		strSql = strSql & " 	and i.isusing = 'Y' "
	    strSql = strSql & " 	and i.isExtusing = 'Y'"
	    strSql = strSql & " 	and i.sellyn='Y'"           '''�Ǹ����� ��ǰ�� ���. // ���� �߰� 2011-11-02
	    strSql = strSql & "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '"&CMALLNAME&"')"	'������ܺ귣��
	    strSql = strSql & "		and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '"&CMALLNAME&"')"		'������ܻ�ǰ
		strSql = strSql & "		and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
	    strSql = strSql & " 	and i.deliverytype not in ('7','6') "   '''���� ��� ���� // ���� �߰� 2011-11-02
'	    strSql = strSql & " 	and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		strSql = strSql & "		and i.cate_large= p.tenCateLarge "
		strSql = strSql & "		and i.cate_mid = p.tenCateMid "
		strSql = strSql & "		and i.cate_small = p.tenCateSmall "
	    strSql = strSql & "		and p.CateKey is Not NULL"   '' �����ڵ�
		strSql = strSql & "		and 'Y' = case	when i.sailyn = 'Y' "
		strSql = strSql & " 					AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 						OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 					) THEN 'Y' "
		strSql = strSql & " 					WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		If FRectItemID <> "" Then
			strSql = strSql & "		and s.itemid in (" & FRectItemID & ")"
		End If
		strSql = strSql & " )"
		strSql = strSql & " and uc.isExtusing <> 'N'"
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid					= rsget("itemid")
				FOneItem.Fitemname					= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid					= rsget("makerid")
				FOneItem.Fsellcash					= rsget("sellcash")
				FOneItem.Forgsellcash				= rsget("orgprice")
				FOneItem.Fsourcearea				= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.FRegdate					= rsget("regdate")
				FOneItem.Fsellyn					= rsget("sellyn")
				FOneItem.Flimityn					= rsget("limityn")
				FOneItem.Flimitno					= rsget("limitno")
				FOneItem.Flimitsold					= rsget("limitsold")
				FOneItem.Fcate_large				= rsget("cate_large")
				FOneItem.Fcate_mid					= rsget("cate_mid")
				FOneItem.Fcate_small				= rsget("cate_small")
				FOneItem.FMakerName					= db2html(rsget("makername"))
				FOneItem.FBrandName					= db2html(rsget("brandname"))
				FOneItem.Foptioncnt					= rsget("optioncnt")
				FOneItem.FBrandNameKor = db2html(rsget("socname_kor"))
			If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
				FOneItem.FMakerName					= FOneItem.FBrandName
			End If
				FOneItem.Fkeywords					= db2html(rsget("keywords"))
				FOneItem.Fbasicimage				= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName				= rsget("regImageName")
				FOneItem.Fmainimage					= rsget("mainimage")
			If IsNULL(FOneItem.FInfoImage) Then
				FOneItem.FInfoImage				= ",,,,"
			End If
				FOneItem.Fordercomment				= db2html(rsget("ordercomment"))
				FOneItem.FItemContent				= db2html(rsget("itemcontent"))
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"��","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea				= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude				= rsget("vatinclude")
			If (rsget("usinghtml") = "N") Then
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,vbcrlf,"<br>")
			End If
				FOneItem.Finterparkdispcategory		= rsget("CateKey")
				FOneItem.Fitemsize					= db2html(rsget("itemsize"))
				FOneItem.Fitemsource				= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate				= rsget("LastUpdate")
				FOneItem.FSellEndDate				= rsget("sellenddate")
				FOneItem.FItemDiv					= rsget("ItemDiv")
				FOneItem.Fisusing					= rsget("isusing")
				FOneItem.FSailYn					= rsget("sailyn")
				FOneItem.FOrgPrice					= rsget("orgprice")
				FOneItem.FdeliveryType				= rsget("deliveryType")
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.Fdeliverfixday				= rsget("deliverfixday")
				FOneItem.Ffreight_min				= rsget("freight_min")
				FOneItem.Ffreight_max				= rsget("freight_max")
				FOneItem.FAdultType 				= rsget("adulttype")
				FOneItem.FOrderMaxNum 				= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getInterparkEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,s.interparkPrdNo, s.mayiParkSellYn "
        strSql = strSql & " ,c.usinghtml, m.CateKey ,s.interparkregdate "
		strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
		strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName, isNULL(s.lastErrStr,'') as lastErrStr, s.mayiparkprice "
		strSql = strSql & " ,(SELECT COUNT(*) as regOptCnt FROM db_item.dbo.tbl_outmall_regedoption as RO WHERE RO.itemid = s.itemid and RO.mallid = 'interpark') as regOptCnt "
		strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM [db_item].[dbo].tbl_interpark_reg_item s, [db_item].[dbo].tbl_item i "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid "
        strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE s.itemid=i.itemid"
		strSql = strSql & addSql
		strSql = strSql & " ORDER BY i.itemid "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid				= rsget("itemid")
				FOneItem.Fitemname				= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid				= rsget("makerid")
				FOneItem.Fbuycash				= rsget("buycash")
				FOneItem.Fsellcash				= rsget("sellcash")
				FOneItem.Forgsellcash			= rsget("orgprice")
				FOneItem.Fsourcearea			= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.Foptioncnt				= rsget("optioncnt")
				FOneItem.FRegdate				= rsget("regdate")
				FOneItem.Fsellyn				= rsget("sellyn")
				FOneItem.Flimityn				= rsget("limityn")
				FOneItem.Flimitno				= rsget("limitno")
				FOneItem.Flimitsold				= rsget("limitsold")
				FOneItem.Fcate_large			= rsget("cate_large")
				FOneItem.Fcate_mid				= rsget("cate_mid")
				FOneItem.Fcate_small			= rsget("cate_small")
				FOneItem.FMakerName				= db2html(rsget("makername"))
				FOneItem.FBrandName				= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor			= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName			= FOneItem.FBrandName
				End If
				FOneItem.Fkeywords				= db2html(rsget("keywords"))
				FOneItem.Fbasicimage			= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName			= rsget("regImageName")
				FOneItem.Fmainimage				= rsget("mainimage")
				FOneItem.Fmainimage2			= rsget("mainimage2")
				If IsNULL(FOneItem.FInfoImage) Then
					FOneItem.FInfoImage			= ",,,,"
				End If
				FOneItem.Fordercomment			= db2html(rsget("ordercomment"))
				FOneItem.FItemContent			= db2html(rsget("itemcontent"))
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"��","")
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent			= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea			= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude			= rsget("vatinclude")
				If (rsget("usinghtml") = "N") Then
					FOneItem.FItemContent		= replace(FOneItem.FItemContent,vbcrlf,"<br>")
				End If
                FOneItem.Finterparkdispcategory	= rsget("CateKey")
				FOneItem.Fitemsize				= db2html(rsget("itemsize"))
				FOneItem.Fitemsource			= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate			= rsget("LastUpdate")
				FOneItem.FSellEndDate			= rsget("sellenddate")
				FOneItem.FItemDiv				= rsget("ItemDiv")
				FOneItem.Fisusing				= rsget("isusing")
				FOneItem.FInterparkPrdNo		= rsget("InterparkPrdNo")
				FOneItem.FmayiParkSellYn		= rsget("mayiParkSellYn")
				FOneItem.FdeliveryType			= rsget("deliveryType")
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.FSailYn				= rsget("sailyn")
                FOneItem.FOrgPrice				= rsget("orgprice")
                FOneItem.Finterparkregdate		= rsget("interparkregdate")
                FOneItem.Fdeliverfixday			= rsget("deliverfixday")
                FOneItem.Ffreight_min			= rsget("freight_min")
                FOneItem.Ffreight_max			= rsget("freight_max")
                FOneItem.FlastErrStr			= rsget("lastErrStr")
                FOneItem.Fmayiparkprice			= rsget("mayiparkprice")
                FOneItem.FregOptCnt				= rsget("regOptCnt")
                FOneItem.FMaySoldOut			= rsget("maySoldOut")
                FOneItem.FbasicimageNm 			= rsget("basicimage")
				FOneItem.FAdultType 			= rsget("adulttype")
				FOneItem.FOrderMaxNum 			= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getInterparkNotRegScheduleOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP 100 i.* "
		strSql = strSql & " ,c.makername, uc.socname_kor, uc.defaultfreeBeasongLimit "
		strSql = strSql & " ,c.keywords, c.ordercomment, c.itemcontent, c.sourcearea "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " ,c.usinghtml, m.CateKey "
		strSql = strSql & " ,isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max "
		strSql = strSql & " ,isNULL(s.regImageName,'') as regImageName "
		strSql = strSql & " FROM [db_item].[dbo].tbl_item i "
		strSql = strSql & " INNER JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		strSql = strSql & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item s on i.itemid = s.itemid "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " where 1=1 "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large <> ''  "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'�ù�(�Ϲ�)
		strSql = strSql & "	and i.itemdiv in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		strSql = strSql & " and i.isusing = 'Y'  "
		strSql = strSql & " and i.isExtusing = 'Y' "
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = '"&CMALLNAME&"') "
		strSql = strSql & " and i.itemid NOT IN (SELECT itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '"&CMALLNAME&"') "
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	and 'Y' = case	when i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & " and i.deliverytype not in ('7','6') "
'		strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000))) "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold> 5 ))) "
		strSql = strSql & " and isnull(s.interParkPrdNo, '') = '' "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CInterparkItem
				FOneItem.Fitemid					= rsget("itemid")
				FOneItem.Fitemname					= LeftB(db2html(rsget("itemname")),255)
				FOneItem.FMakerid					= rsget("makerid")
				FOneItem.Fsellcash					= rsget("sellcash")
				FOneItem.Forgsellcash				= rsget("orgprice")
				FOneItem.Fsourcearea				= LeftB(db2html(rsget("sourcearea")),64)
				FOneItem.FRegdate					= rsget("regdate")
				FOneItem.Fsellyn					= rsget("sellyn")
				FOneItem.Flimityn					= rsget("limityn")
				FOneItem.Flimitno					= rsget("limitno")
				FOneItem.Flimitsold					= rsget("limitsold")
				FOneItem.Fcate_large				= rsget("cate_large")
				FOneItem.Fcate_mid					= rsget("cate_mid")
				FOneItem.Fcate_small				= rsget("cate_small")
				FOneItem.FMakerName					= db2html(rsget("makername"))
				FOneItem.FBrandName					= db2html(rsget("brandname"))
				FOneItem.Foptioncnt					= rsget("optioncnt")
				FOneItem.FBrandNameKor = db2html(rsget("socname_kor"))
			If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
				FOneItem.FMakerName					= FOneItem.FBrandName
			End If
				FOneItem.Fkeywords					= db2html(rsget("keywords"))
				FOneItem.Fbasicimage				= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FregImageName				= rsget("regImageName")
				FOneItem.Fmainimage					= rsget("mainimage")
			If IsNULL(FOneItem.FInfoImage) Then
				FOneItem.FInfoImage				= ",,,,"
			End If
				FOneItem.Fordercomment				= db2html(rsget("ordercomment"))
				FOneItem.FItemContent				= db2html(rsget("itemcontent"))
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"��","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,"","")
				FOneItem.Fsourcearea				= db2html(rsget("sourcearea"))
				FOneItem.Fvatinclude				= rsget("vatinclude")
			If (rsget("usinghtml") = "N") Then
				FOneItem.FItemContent				= replace(FOneItem.FItemContent,vbcrlf,"<br>")
			End If

				FOneItem.Finterparkdispcategory		= rsget("CateKey")
				FOneItem.Fitemsize					= db2html(rsget("itemsize"))
				FOneItem.Fitemsource				= db2html(rsget("itemsource"))
				FOneItem.FLastUpdate				= rsget("LastUpdate")
				FOneItem.FSellEndDate				= rsget("sellenddate")
				FOneItem.FItemDiv					= rsget("ItemDiv")
				FOneItem.Fisusing					= rsget("isusing")
				FOneItem.FSailYn					= rsget("sailyn")
				FOneItem.FOrgPrice					= rsget("orgprice")
		'2012-11-09 ���� ����(���̾ ��ǰ�̸� ������
'		2017-11-27 ���� ���� / ���Ƹ� �븮�� ��û (���̾ ��ǰ�� 30000�� �̻� ����, �̸��� �� 2500���� ������û
'			If (IsNull(rsget("DyItemid")) = "False" and CLng(rsget("sellcash")) > 13000) AND ((rsget("cate_large") = "010") AND (rsget("cate_mid") = "010") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "020") OR (rsget("cate_large") = "010") AND (rsget("cate_mid") = "030") ) Then
'				FOneItem.FdeliveryType				= "4"
'			Else
				FOneItem.FdeliveryType				= rsget("deliveryType")
'			End If
				FOneItem.FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FOneItem.Fdeliverfixday				= rsget("deliverfixday")
				FOneItem.Ffreight_min				= rsget("freight_min")
				FOneItem.Ffreight_max				= rsget("freight_max")
				FOneItem.FAdultType 				= rsget("adulttype")
		End If
		rsget.Close
	End Sub
End Class

Function getInterparkPrdno(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNULL(interparkPrdNo,'') as interparkPrdNo "&VbCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_interpark_reg_Item "&VbCRLF
	sqlStr = sqlStr & " WHERE itemid="&iitemid
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		retVal = rsget("interparkPrdNo")
	End if
	rsget.Close
	If IsNULL(retVal) Then retVal=""
	getInterparkPrdno = retVal
End Function

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function
%>
