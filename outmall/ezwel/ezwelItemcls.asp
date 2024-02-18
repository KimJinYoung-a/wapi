<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "ezwel"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CEzwelMARGIN = 10									'��������� ���� 10%
CONST cspCd		= "10040413"							'CP��ü�ڵ�(������ �߱�)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'�����ڵ�(������ �߱�)
CONST cspDlvrId	= "10040413"							'���ó�ڵ�

Class CEzwelItem
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
	Public FezwelStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FdepthCode
	Public FbasicimageNm
	Public FezwelGoodNo
	Public Fezwelprice
	Public FezwelSellYn
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
    Public FsafetyNum
    Public FmaySoldOut
	Public FAdultType

    Public Fregitemname
    Public FregImageName
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
		End If
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	public Function getKeywords()
		Dim strRst
		strRst = FKeywords
		strRst = replace(strRst, "�α�", "")
		strRst = replace(strRst, "��ġ", "")
		strRst = replace(strRst, "�����ġ", "")
		If strRst = "" Then
			strRst = "�ٹ�����"
		End If
		getKeywords = Server.URLEncode(strRst)
	End Function

	public Function getNewKeywords()
		Dim strRst
		strRst = FKeywords
		strRst = replace(strRst, "�α�", "")
		strRst = replace(strRst, "��ġ", "")
		strRst = replace(strRst, "�����ġ", "")
		If strRst = "" Then
			strRst = "�ٹ�����"
		End If
		getNewKeywords = strRst
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt
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
				If (FLimitYN <> "Y") Then optLimit = 999

				If (optLimit <> 0) Then
					limitYCnt =  limitYCnt + 1
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If limitYCnt = 0 Then
			IsMayLimitSoldout = "Y"
		Else
			IsMayLimitSoldout = "N"
		End If
	End Function

	Function getiszeroWonSoldOut(iitemid)
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
					If goptlimitno - goptlimitsold > CMAXLIMITSELL Then
						i = i + 1
					End If
					rsget.MoveNext
				Loop

				If i = 0 Then		'0�� �ɼ��� ��� 5�� ���ϸ� ǰ��
					getiszeroWonSoldOut = "Y"
				Else
					getiszeroWonSoldOut = "N"
				End If
			Else
				getiszeroWonSoldOut = "Y"
			End If
			rsget.Close
		End If
	End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin, sqlStr, specialPrice
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

		If specialPrice <> "" Then
			fngetMustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				fngetMustPrice = Forgprice
			Else
				fngetMustPrice = FSellCash
			End If
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice
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

		If specialPrice <> "" Then
			MustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
		End If
	End Function

	'// Ezwel �Ǹſ��� ��ȯ
	Public Function getEzwelSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEzwelSellYn = "Y"
			Else
				getEzwelSellYn = "N"
			End If
		Else
			getEzwelSellYn = "N"
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

	Public Function getLimitEzwelEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 10000 Then
				ret = 10000
			End If
		Else
			ret = 10000
		End If

		If (ret < 1) Then ret = 0
		getLimitEzwelEa = ret
	End Function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

    Function getEzwelAddSuplyPrice(addprice)
		getEzwelAddSuplyPrice = CLNG((addprice)  * (100-CEzwelMARGIN) / 100)
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
'				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
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

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(�ٹ�), 4,5(����)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'��ü����
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	Public Function getBrandCode(v)
		Dim strSql
		strSql = strSql & " SELECT TOP 1 brandCd "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ezwel_brandList] "
		'strSql = strSql & " WHERE brandNm like '%"& html2db(v) &"%' "
		strSql = strSql & " WHERE brandNm = '"& html2db(v) &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getBrandCode = rsget("brandCd")
		Else
			getBrandCode = "143289"
		End If
		rsget.Close
	End Function

	Public Function getMafcCode(v)
		Dim strSql
		strSql = strSql & " SELECT TOP 1 mafcCd "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ezwel_mafcList] "
		'strSql = strSql & " WHERE mafcNm like '%"& html2db(v) &"%' "
		strSql = strSql & " WHERE mafcNm = '"& html2db(v) &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getMafcCode = rsget("mafcCd")
		Else
			getMafcCode = "184231"
		End If
		rsget.Close
	End Function

	'��ǰ���� �Ķ���� ����
	Public Function getEzwelItemContParam()
		Dim strRst, strSQL,strRst2
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
		Fitemcontent = rpTxt(Fitemcontent)

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")

		strRst = replace(replace(strRst, "<script", ""), "</script>", "")
		getEzwelItemContParam = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','ezwel') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst2 = rpTxt(rsget("textVal"))
		'response.end
			strRst = ("<div align=""center"">")
			strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
			strRst = strRst & strRst2
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
			strRst = strRst & ("</div>")
			getEzwelItemContParam = strRst
		End If
		rsget.Close

	End Function

	'��ǰ���� �Ķ���� ����
	Public Function getEzwelNewItemContParam()
		Dim strRst, strSQL,strRst2
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
		strRst = strRst & ("</div>")

		strRst = replace(replace(strRst, "<script", ""), "</script>", "")
		getEzwelNewItemContParam = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','ezwel') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst2 = rsget("textVal")
		'response.end
			strRst = ("<div align=""center"">")
			strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_ezwel.jpg""></center></p><br>")
			strRst = strRst & strRst2
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ezwel.jpg"">")
			strRst = strRst & ("</div>")
			getEzwelNewItemContParam = strRst
		End If
		rsget.Close
	End Function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����
	Public Function getEzwelAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If

		strRst = strRst &"	<imgPath><![CDATA["&FbasicImage&"]]></imgPath>"		'�����̹������ | ex)http://www.ezwel.com/img/goods1.gif
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'�߰��̹������1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"	<imgPath"&i&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"]]></imgPath"&i&">"
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next

		End If
		rsget.Close
		getEzwelAddImageParam = strRst
	End Function

	Public Function getEzwelNewAddImageParam(obj)
		Dim strSQL, i
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If
		obj("imgPath") = FbasicImage						'�����̹������

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		'�߰��̹������1~3
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					obj("imgPath"&i&"") = "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")		'�߰��̹������1~3
				End If
				rsget.MoveNext
				If i >= 3 Then Exit For
			Next
		End If
		rsget.Close
	End Function

	'��ǰǰ������
    public function getEzwelItemInfoCd()
		Dim buf1, buf2, buf3, strSQL, mallinfoCd, infoContent, mallinfodiv
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (IC.safetyyn <> 'Y' ) THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '������������' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '�����ŷ�����ȸ ���(�Һ��ں����ذ����)�� �ǰ��Ͽ� ������ �帳�ϴ�.' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='J' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSql = strSql & "		 WHEN c.infotype='K' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		'strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035'  " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '�ٹ����� ��ǰ���� / Q&A �ۼ�'  " & vbcrlf
		strSql = strSql & "		 WHEN LEN( isNull(F.infocontent, '')) < 2 THEN '��ǰ �� ����' " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ezwel' and IC.itemid='"&FItemid&"'  " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		''mallinfodiv = "10" & rsget("mallinfodiv")  '' �̵� eastone 2016/08/17
		If Not(rsget.EOF or rsget.BOF) then
		    mallinfodiv = "10" & rsget("mallinfodiv")
			If mallinfodiv = "1047" Then
				mallinfodiv = "1039"
			ElseIf mallinfodiv = "1048" Then
				mallinfodiv = "1040"
			End If

			buf1 = "<goodsGrpCd>"&mallinfodiv&"</goodsGrpCd>"		'##*��ǰ��� �ڵ� | ����÷��
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")

				If FMakerid = "indigoshop" and (rsget("infocd") = "35002") Then
					infoContent = ".."
				end if

				If rsget("infocontent") = "" or isnull(infocontent) Then
					infoContent = "�������� ����"
				End If

				buf2 = buf2 & " 		<arrLayoutDesc><![CDATA["& Server.URLEncode(infoContent) &"]]></arrLayoutDesc>"
				buf2 = buf2 & " 		<arrLayoutSeq>"&mallinfoCd&"</arrLayoutSeq>"
				rsget.MoveNext
			Loop
			buf3 = buf1 & buf2
		End If
		rsget.Close
		getEzwelItemInfoCd = buf3
	End Function

	Public Function getEzwelItemNewInfoCd(obj)
		Dim strSQL, mallinfoCd, infoContent, mallinfodiv, i
		strSQL = "EXEC [db_etcmall].[dbo].[usp_API_Ezwel_InfoCodeMap_Get] " & FItemID
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
		    mallinfodiv = "10" & rsget("mallinfodiv")
			If mallinfodiv = "1047" Then
				mallinfodiv = "1039"
			ElseIf mallinfodiv = "1048" Then
				mallinfodiv = "1040"
			End If
			obj("goodsGrpCd") = mallinfodiv							'#��ǰ��� �ڵ�
			Set obj("arrLayoutDesc") = jsArray()					'#��ǰ��� ����
			Set obj("arrLayoutSeq") = jsArray()						'#��ǰ��� �׸� ����
			i = 0
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")

				If FMakerid = "indigoshop" and (rsget("infocd") = "35002") Then
					infoContent = ".."
				end if

				If rsget("infocontent") = "" or isnull(infocontent) Then
					infoContent = "�������� ����"
				End If
				obj("arrLayoutDesc")(i) = infoContent
				obj("arrLayoutSeq")(i) = mallinfoCd
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Function

   Public Function getEzwelOptionParam()
		Dim strSql, strRst, i, optLimit, sellOptcnt
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option with (nolock) "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				sellOptcnt = rsget("cnt")
			rsget.Close

			If sellOptcnt > 0 Then
				strSql = ""
				strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				optDc = ""
				optLimit = ""
				If FVatInclude = "N" Then
					optTaxCk = "N"
				Else
					optTaxCk = "Y"
				End If

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 ���������� ��� Y�� ���� �ǹǷ�
						optUsingCk = "Y"
						optDc = optDc & Server.URLEncode(rpTxt(db2Html(replace(rsget("optionname"), ":", ""))))

						itemsu = itemsu & optLimit
						addprice = addprice & rsget("optaddprice")
						addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
						optTax = optTax & optTaxCk
						optUsing = optUsing & optUsingCk

						rsget.MoveNext
						If Not(rsget.EOF) Then
							optDc	= optDc & "|"
							itemsu = itemsu & "|"
							addprice = addprice & "|"
							addbuyprice = addbuyprice & "|"
							optTax	= optTax & "|"
							optUsing = optUsing & "|"
						End If
					Loop
				End If
				rsget.Close
				buf = buf & "		<useYn>Y</useYn>"												'��ǰ�ɼǻ�뿩�� | �ɼ��� �������(Y) �������(N)
				buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("����")&"</arrOptionCdNm>"	'��ǰ�ɼǸ�
				buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'��ǰ�ɼ� ����
				buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'�ɼǺ��� ���� ��뿩�� | Y:N
				buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(�ɼ��� �����ϴ� ��츸) | ��ǰ�ɼ� ���� | Default: 10000
				buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'��ǰ�ɼ��߰�����
				buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'���ް�
				buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'�������� | ����(Y), �鼼(N), ����(���� 0)
				buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'�ɼ� �󼼺��� ���� ��뿩�� |||    Y|Y|Y:N|N:N
			Else
				strSql = ""
				strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				optDc = ""
				optLimit = ""
				If FVatInclude = "N" Then
					optTaxCk = "N"
				Else
					optTaxCk = "Y"
				End If

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = 999   ''2013/06/12 ���������� ��� Y�� ���� �ǹǷ�
						optUsingCk = "N"
						optDc = optDc & Server.URLEncode(rpTxt(db2Html(replace(rsget("optionname"), ":", ""))))

						itemsu = itemsu & optLimit
						addprice = addprice & rsget("optaddprice")
						addbuyprice = addbuyprice & getEzwelAddSuplyPrice(rsget("optaddprice"))
						optTax = optTax & optTaxCk
						optUsing = optUsing & optUsingCk

						rsget.MoveNext
						If Not(rsget.EOF) Then
							optDc	= optDc & "|"
							itemsu = itemsu & "|"
							addprice = addprice & "|"
							addbuyprice = addbuyprice & "|"
							optTax	= optTax & "|"
							optUsing = optUsing & "|"
						End If
					Loop
				End If
				rsget.Close
				buf = buf & "		<useYn>Y</useYn>"												'��ǰ�ɼǻ�뿩�� | �ɼ��� �������(Y) �������(N)
				buf = buf & "		<arrOptionCdNm>"&Server.URLEncode("����")&"</arrOptionCdNm>"	'��ǰ�ɼǸ�
				buf = buf & "		<arrOptionContent>"&optDc&"</arrOptionContent>"					'��ǰ�ɼ� ����
				buf = buf & "		<arrOptionUseYn>Y</arrOptionUseYn>"								'�ɼǺ��� ���� ��뿩�� | Y:N
				buf = buf & "		<arrOptionAddAmt>"&itemsu&"</arrOptionAddAmt>"					'*(�ɼ��� �����ϴ� ��츸) | ��ǰ�ɼ� ���� | Default: 10000
				buf = buf & "		<arrOptionAddPrice>"&addprice&"</arrOptionAddPrice>"			'��ǰ�ɼ��߰�����
				buf = buf & "		<arrOptionAddBuyPrice>"&addbuyprice&"</arrOptionAddBuyPrice>"	'���ް�
				buf = buf & "		<arrOptionAddTaxYn>"&optTax&"</arrOptionAddTaxYn>"				'�������� | ����(Y), �鼼(N), ����(���� 0)
				buf = buf & "		<arrOptionFullUseYn>"&optUsing&"</arrOptionFullUseYn>"			'�ɼ� �󼼺��� ���� ��뿩�� |||    Y|Y|Y:N|N:N
			End If
		Else
			buf = buf & "		<useYn>N</useYn>"												'��ǰ�ɼǻ�뿩�� | �ɼ��� �������(Y) �������(N)
		End If
		getEzwelOptionParam = buf
    End Function

	Public Function getEzwelNewOptionParam(obj)
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing
'FoptionCnt = 0
    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, itemoption "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				obj("useYn") = "Y"										'��ǰ�ɼǻ�뿩��
				obj("optType") = "1001"									'��ǰ�ɼ����� | �ܵ���(1001), ������(1002)
				Set obj("optionContentList") = jsArray()				'��ǰ�ɼǸ��
					Set obj("optionContentList")(0) = jsObject()
						obj("optionContentList")(0)("optionCdNm") = "����"
				Set obj("optionFullContentList") = jsArray()

				i = 0
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999

					Set obj("optionFullContentList")(i) = jsObject()
						obj("optionFullContentList")(i)("optionCdNm") = "����"						'��ǰ�ɼǸ�
						obj("optionFullContentList")(i)("optionContent1") = db2Html(rsget("optionname"))		'�ɼǳ���1
						obj("optionFullContentList")(i)("optionAddAmt") = optLimit 					'�ɼǼ���
						obj("optionFullContentList")(i)("optionAddBuyPrice") = getEzwelAddSuplyPrice(rsget("optaddprice"))	'�ɼǸ��԰�
						obj("optionFullContentList")(i)("optionAddPrice") =  rsget("optaddprice")	'�ɼ��߰�����
						obj("optionFullContentList")(i)("useYn") = "Y"								'�ɼǻ󼼻�뿩��
						obj("optionFullContentList")(i)("imgPath") = ""								'�ɼǽ�����̹���
						obj("optionFullContentList")(i)("imgDispYn") = "N"							'�ɼ��̹������⿩��
						obj("optionFullContentList")(i)("sortNo") = i + 1							'�ɼ����ļ���
						obj("optionFullContentList")(i)("imgDtlPath") = ""							'�ɼǻ��̹���
						obj("optionFullContentList")(i)("cspOptionFullNum") = rsget("itemoption")	'��ü�ɼǻ��ڵ�
					rsget.MoveNext
					i = i + 1
				Loop
			End If
			rsget.Close
		Else
			obj("useYn") = "N"					'��ǰ�ɼǻ�뿩��
		End If
	End Function

	Public Function getEzwelCertParameter(obj)
		Dim strSql, safetyDiv, certNum, certOrganName, modelName, certDate, isRegCert
		Dim authType, authNum, certDiv

		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isNull(t.certNum, '') as certNum, isNull(f.modelName, '') as modelName, isNull(f.certDate, '') as certDate, isNull(f.certOrganName, '') as certOrganName, isNull(f.certDiv, '') as certDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
			certOrganName = rsget("certOrganName")
			modelName	= rsget("modelName")
			certDate	= rsget("certDate")
			certDiv		= rsget("certDiv")
			isRegCert	= "Y"
		Else
			isRegCert	= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			Select Case safetyDiv
				Case "10", "40", "70"
					authType		= "1001"
					authNum			= certNum
				Case "20", "50", "80"
					authType		= "1002"
					authNum			= certNum
				Case "30", "60", "90"
					authType		= "1003"
					authNum			= ""
			End Select

			If len(certDate) = 8 Then
				certDate = Left(certDate,4)&"-"&Mid(certDate,5,2)&"-"&Mid(certDate,7,2)
			Else
				certDate = ""
			End If

			obj("safeAuthYn") = "Y"											'������� ����
			obj("authType") = authType										'������� ǰ�� | 1001:��������/1002:����Ȯ��/1003:���������Լ�Ȯ��
			obj("authNum") = authNum										'������ȣ
			obj("authDt") = certDate										'�������� ���� | ex)20220404
			obj("authDiv") = certDiv										'�������� �׸�
			obj("authOrganNm") = certOrganName								'�������� ���
		Else
			obj("safeAuthYn") = "N"											'������� ����
			obj("authType") = ""											'������� ǰ�� | 1001:��������/1002:����Ȯ��/1003:���������Լ�Ȯ��
			obj("authNum") = ""												'������ȣ
			obj("authDt") = ""												'�������� ���� | ex)20220404
			obj("authDiv") = ""												'�������� �׸�
			obj("authOrganNm") = ""											'�������� ���
		End If	
	End Function
	
	Public Function getEzwelDlvrCode(iDepthCode)
		Select Case iDepthCode
			Case "45020518", "45020519", "45110106", "45110105", "45110101", "45110214", "45110212", "45110213", "45110210", "45110211", "45110207", "45110201", "45110205", "45110203", "45110202", "45110215", "70040114"	getEzwelDlvrCode = "1003"
			Case Else
				If FItemdiv = "06" OR FItemdiv = "16" Then
					getEzwelDlvrCode = "1003"
				Else
					getEzwelDlvrCode = "1001"
				End If
		End Select
	End Function

	'��ǰ���/���� XML ����
	Public Function getEzwelItemRegXML(ezwelMethod, ichkXML)
		Dim strRst
		Dim EzwelStatus
		Select Case ezwelMethod
			Case "Reg"			EzwelStatus = "1001"
			Case "SellY"		EzwelStatus = "1002"
			Case "SellN"		EzwelStatus = "1005"
			Case "MustNotOpt"	EzwelStatus = "1005"
		End Select
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "	<dataSet>"
		strRst = strRst & "		<cspCd>"&cspCd&"</cspCd>"					'##*CP ��ü�ڵ� | ������ �߱�(������)
		If ezwelMethod <> "Reg" Then
		strRst = strRst & "		<goodsCd>"&FEzwelGoodno&"</goodsCd>"		'##*���� �����ϸ� ���� �������� ������ �Է� | ��ǰ�ڵ� | ������ ��ǰ�ڵ�
		End If
		strRst = strRst & "		<cspGoodsCd>"&FItemid&"</cspGoodsCd>"		'##��ü��ǰ�ڵ�
		strRst = strRst & "		<goodsNm><![CDATA["&Server.URLEncode(Trim(getItemNameFormat))&"]]></goodsNm>"			'##*��ǰ��
		strRst = strRst & "		<taxYn>"&CHKIIF(FVatInclude="N","N","Y")&"</taxYn>"										'##*�������� | ����(Y), �鼼(N), ����(���� 0)
'		If EzwelStatus <> "1002" Then
			strRst = strRst & "		<goodsStatus>"&EzwelStatus&"</goodsStatus>"												'##��ǰ���� | ���(1001), �Ǹ���(1002), �Ǹ�����(1005), ����(1006), �Ͻ�ǰ��(1004) 2017-11-13 ������..1005�� �Ұ�� MD ���ι޾ƾ� �Ǹ������� �����
'		End If
		strRst = strRst & "		<dlvrPrice>"&CHKIIF(IsFreeBeasong=False,"3000","0")&"</dlvrPrice>"						'##��۰���
		strRst = strRst & "		<dlvrPriceApplYn>"&CHKIIF(IsFreeBeasong=True,"Y","P")&"</dlvrPriceApplYn>"				'##*����/������/���� | ����: Y/ �Һ��ںδ�:N /���Ҹ�: A /��������:P
		strRst = strRst & "		<realSalePrice>"&Clng(GetEzwel10wonDown(MustPrice/10)*10)&"</realSalePrice>"			'##*�ǸŰ�
		strRst = strRst & "		<normalSalePrice>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</normalSalePrice>"			'##*����(����)��
		strRst = strRst & "		<brandNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("��ǰ���� ����"),Server.URLEncode(rpTxt(Fmakername)))&"]]></brandNm>"	'##�귣���
		strRst = strRst & "		<buyPrice>"&GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))&"</buyPrice>"	'##*���ް�(���԰�)
		strRst = strRst & "		<modelNum>"&FItemid&"</modelNum>"														'��ǰ��
		strRst = strRst & "		<orginNm><![CDATA["&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),Server.URLEncode("��ǰ���� ����"),Server.URLEncode(Fsourcearea))&"]]></orginNm>"	'##������
		strRst = strRst & "		<mafcNm><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),Server.URLEncode("��ǰ���� ����"),Server.URLEncode(rpTxt(Fmakername)))&"]]></mafcNm>"		'##������
		strRst = strRst & "		<enterAmt>"&getLimitEzwelEa()&"</enterAmt>"						'##*�԰���� | Default: 10000
		strRst = strRst & "		<cspDlvrId>"&cspDlvrId&"</cspDlvrId>"			'##�����ID | ������ �߱�(������)
		strRst = strRst & "		<goodsDesc><![CDATA["&Server.URLEncode(getEzwelItemContParam())&"]]></goodsDesc>"		'##��ǰ����
		If (ezwelMethod <> "Reg") Then		'2014-12-02 ������ �߰� | �̹��� ���� �ð� �����ɸ�
			If isImageChanged Then
				strRst = strRst & getEzwelAddImageParam()
			End If
		Else
			strRst = strRst & getEzwelAddImageParam()
		End If
		strRst = strRst & "		<ctgCd>"&FDepthCode&"</ctgCd>"					'##*����ī�װ� | ����÷��
		strRst = strRst & "		<dispCtgCd>"&FDepthCode&"</dispCtgCd>"			'##*���� ī�װ� | ����÷��
		strRst = strRst & getEzwelItemInfoCd()									'##��ǰ����������� �ʵ����� | ��ǰ�������� ��ø� ���� �ʵ�����
		If ezwelMethod = "MustNotOpt" Then
			strRst = strRst & "	<useYn>N</useYn>"
		Else
			strRst = strRst & getEzwelOptionParam()
		End If

		strRst = strRst & "		<arrIconCd>1008</arrIconCd>"					'������ | ���� = 1008 / ������ = 1010 / ���κ��� = 1007	'2018-08-23 ������ 1008��û
		strRst = strRst & "		<marginRate>"&CEzwelMARGIN&"</marginRate>"		'##���ƴ븮�� 10%��� �亯 | *������ | 9.0
		strRst = strRst & "		<dlvrForm>"&getEzwelDlvrCode(FDepthCode)&"</dlvrForm>"			'������� | 1001 : �Ϲ��ù�, 1002 : ��ü���, 1003 : �ֹ�����, 1004 : ��ġ��ǰ
		strRst = strRst & "		<keyword><![CDATA["&RightCommaDel(Trim(getKeywords()))&"]]></keyword>"			'�˻�Ű���� | ���� Ű���� �Է°��� (,)�� ���� ex)����,�ؿ�����,����귣��
		strRst = strRst & "		<unitOrderQty>"& FOrderMaxNum &"</unitOrderQty>"	'�δ籸�ż��� | 1ȸ�� ������ �� �ִ� ���� ���� * ���� ������ �ʰų� 0�ΰ�� �������� ����
		strRst = strRst & "</dataSet>"
		getEzwelItemRegXML = strRst
If (session("ssBctID")="kjy8517") Then
		response.write replace(strRst, "?xml", "?AAAAAl")
'		response.end
End If
	End Function

	'��ǰ���/���� Json ����
	Public Function getEzwelItemRegJson(v)
		Dim obj
		Set obj = jsObject()
			If v = "EDIT" Then
				obj("goodsCd") = FEzwelGoodNo
			End If

			If application("Svr_Info")="Dev" Then
				FDepthCode = "70040114"
			End If
			obj("cspGoodsCd") = FItemid										'��ü��ǰ�ڵ�
			obj("goodsNm") = getItemNameFormat()							'#��ǰ��
			obj("taxYn") = CHKIIF(FVatInclude="N","N","Y")					'#�������� | ����(Y), �鼼(N), ����(���� 0)
			obj("goodsStatus") = "1001"										'��ǰ���� | ���(1001), �Ǹ���(1002), �Ǹ�����(1005), ����(1006)
			obj("dlvrPrice") = CHKIIF(IsFreeBeasong=False,"3000","0")		'��۰���
			obj("addJejuDlvrPrice") = "3000"								'�߰���ۺ�(����)
			obj("addSanganDlvrPrice") = "3000"								'�߰���ۺ�(�����갣)
			obj("dlvrPriceApplYn") = CHKIIF(IsFreeBeasong=True,"Y","P")		'#����./������/���� | ����: Y / �Һ��ںδ�:N /���Ҹ�: A /��������:P /����(����������):C
			obj("realSalePrice") = Clng(GetEzwel10wonDown(MustPrice/10)*10)	'#�ǸŰ�
			obj("normalSalePrice") = Clng(GetRaiseValue(ForgPrice/10)*10)	'#����(����)��
			obj("brandNm") = chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����", Fmakername)	'�귣���
			obj("brandCd") = getBrandCode(Fmakername)						'#�귣���ڵ�
			obj("buyPrice") = GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))							'#���ް�(���԰�)
			obj("modelNum") = FItemid										'��ǰ��
			obj("orginNm") = chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea), "��ǰ���� ����", Fsourcearea)		'������
			obj("mafcNm") = chkIIF(trim(Fmakername)="" or isNull(Fmakername), "��ǰ���� ����", Fmakername)	'������
			obj("mafcCd") = getMafcCode(Fmakername)							'#�������ڵ�
			obj("enterAmt") = getLimitEzwelEa()								'#�԰����
			obj("cspDlvrId") = cspDlvrId									'�����ID
			obj("saleStartDt") = Replace(Date(), "-", "")					'�ǸŽ�������
			obj("saleEndDt") = "20991231"									'�Ǹ���������
			obj("goodsDesc") = getEzwelNewItemContParam()					'��ǰ����
			Call getEzwelNewAddImageParam(obj)
			obj("ctgCd") = FDepthCode										'#���� ī�װ�
			obj("dispCtgCd") = FDepthCode									'#���� ī�װ�
			Call getEzwelItemNewInfoCd(obj)
			Call getEzwelNewOptionParam(obj)
'			obj("iconCd") = ""												'��ǰ������ | ī�װ��� ��ǰ�� ��� 1011(���)/1012(�õ�)/1013(����)/1014(�ش����)
			Set obj("arrIconCd") = jsArray()								'������ | ���� = 1008 / ������(���κ���) = 1007
				obj("arrIconCd")(0) = "1008"
			obj("marginRate") = CEzwelMARGIN								'#������
			obj("dlvrForm") = getEzwelDlvrCode(FDepthCode)					'#��ǰ���� | 1001:�Ϲ��ù�,1002:��ü���,1003:�ֹ�����,1004:��ġ��ǰ,1005:�ؿ������,1006:�Ǹ������Ĺ���,1007:����/�õ���ǰ,1008:�ż���ǰ
			obj("exchgPrice") = 3000										'��ȯ ��ۺ�
			obj("returnPrice") = 3000										'��ǰ ��ۺ�
'			obj("bndlNonChgReturnYn") = ""									'������ȯ/��ǰ�Ұ� | Y/N
			obj("keyword") = RightCommaDel(Trim(getNewKeywords()))			'�˻�Ű���� | ���� Ű�����Է°��� (,)�� ���� ex)����, �ؿ�����, ����귣��
'			obj("shortDesc") = ""											'��ǰȫ������
			obj("policyNo") = "10744781"									'#�߼���å ������ | CP�߼���å �������� ����
'			obj("imgPath640") = ""											'������̹������(640*320)
			obj("dlvrFreeYn") = "Y"											'���Ǻ� ������ ���� | Y:���, N:�̻�� *�����ID(cspDlvrId)�� ������ ���Ǻ� ������ ���ΰ� N�ΰ�� ������ N���� ���(���/��ǰ�� ���/���� API ����)
			obj("unitOrderQty") = FOrderMaxNum								'#1ȸ��  ���ż��� | 1ȸ�� ������ �� �ִ� ���� ���� * ���� ������ �ʰų� 0�ΰ�� �������� ����
			obj("idUnitOrderQty") = 0										'�δ籸�ż���(�⵵) | 1�⿡ �� ���̵� �� ������ �� �ִ� ���� ���� * ���� ������ �ʰų� 0�� ��� �������� ����
'			obj("minPriceYn") = ""											'������Ȯ�� | ���������� ���:Y/������:N/��ǰ���:D/�̸�Ī:M
'			obj("minPriceUrl") = ""											'������ ��ũ | ����)https://search.shopping.naver.com/detail/detail.nhn?nv_mid=xxxx
			obj("exceptBndlDlvrYn") = "N"									'������� ���ܿ��� | Y:������� �Ұ�, N:������� ����
			obj("goodsType") = "1001"										'��ǰ���� | 1001: �Ϲݻ�ǰ, 1002: �޴�����ǰ
'			obj("arrQuotaAmt") = ""											'#�Һο���
'			obj("arrSaleTypeSp") = ""										'�Ǹű��� | 1001:�ű԰���, 1002:��ȣ�̵�, 1003:��⺯��
'			obj("arrGuide") = ""											'#�ű԰��Ծȳ��޽���
'			obj("arrSaleStopDesc") = ""										'#�Ǹ�����ȳ��޽���
'			obj("arrJoinAmt1") = ""											'���Ժ� -���� ����û�� �޽���
'			obj("arrJoinAmt2") = ""											'���Ժ� - �������� �볳 �޽���
'			obj("arrJoinAmt3") = ""											'���Ժ� - ����, �����/���������� �޽���
'			obj("arrJoinAmt4") = ""											'���Ժ� - ����, �簡�� �޽���
'			obj("arrJoinYn1") = ""											'���Ժ� -���� ����û�� �޽��� ��뿩��
'			obj("arrJoinYn2") = ""											'���Ժ� - �������� �볳 �޽��� ��뿩��
'			obj("arrJoinYn3") = ""											'���Ժ� - ����, �����/���������� �޽��� ��뿩��
'			obj("arrJoinYn4") = ""											'���Ժ� - ����, �簡�� �޽��� ��뿩��
'			obj("arrUsimAmt1") = ""											'���ɺ� - ���� �޽���
'			obj("arrUsimAmt2") = ""											'���ɺ� - ���� �޽���
'			obj("arrUsimYn1") = ""											'���ɺ� - ���� �޽��� ��뿩��
'			obj("arrUsimYn2") = ""											'���ɺ� - ���� �޽��� ��뿩��
'			obj("arrTerm1") = ""											'���� - �Ǹ����� �޽���
'			obj("arrTemr2") = ""											'���� - �ΰ����� �޽���
'			obj("arrSaleStopYn") = ""										'�Ǹ����� ����
'			obj("arrQuotaMonth") = ""										'�Һΰ��� �� | 24: 24����, 30:30����
'			obj("arrMsg") = ""												'#�Һΰ��� �޽���
'			obj("arrPrepayAmt") = ""										'#�������ݾ�
'			obj("saleTypeUrl1") = ""										'�ű԰��� ���� URL
'			obj("saleTypeUrl2") = ""										'��ȣ�̵� ���� URL
'			obj("saleTypeUrl3") = ""										'��⺯�� ���� URL
'			obj("noticeNm") = ""											'���ǻ��׸�
'			obj("noticeDesc") = ""											'���ǻ��׳���
'			obj("noticeOrderNo") = ""										'���ǻ������ļ���
'			obj("arrPriceCd") = ""											'������ڵ�
'			obj("arrMobileUseYn") = ""										'��뿩��
'			obj("arrMbDcCd") = ""											'�ܸ��������ڵ�
'			obj("arrFixDcCd") = ""											'���������ڵ�
'			obj("arrQuotaMonthPrice") = ""									'�Һΰ����� | 24: 24����, 30:30����
			obj("adultAuthYn") = IsAdultItem()								'���� �������� �ʿ��ǰ������� | Y�Ͻ� ����� ȭ�鿡�� 19�� �̻� ���������� ��ģ ��, ��ǰ���Ű���
			Call getEzwelCertParameter(obj)
			getEzwelItemRegJson = obj.jsString
		Set obj = nothing
	End Function

	'��ǰ���ݺ��� Json ����
	Public Function getEzwelItemPriceJson()
		Dim obj
		Set obj = jsObject()
			obj("goodsCd") = FEzwelGoodno
			obj("realSalePrice") = Clng(GetEzwel10wonDown(MustPrice/10)*10)	'#�ǸŰ�
			obj("buyPrice") = GetEzwelBuyPrice(Clng(GetEzwel10wonDown(MustPrice/10)*10))	'#���ް�(���԰�)
			getEzwelItemPriceJson = obj.jsString
		Set obj = nothing
	End Function

	'��ǰ�ɼǺ��� Json ����
	Public Function getEzwelItemOptionJson()
		Dim obj
		Dim strSql, strRst, i, optLimit
    	Dim buf, optDc, itemsu, addprice, addbuyprice, optTaxCk, optTax, optUsingCk, optUsing
'FoptionCnt = 0
		Set obj = jsObject()
			obj("goodsCd") = FEzwelGoodno

    	buf = ""
		If FoptionCnt>0 then
			strSql = ""
			strSql = strSql &  "SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, itemoption "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			optDc = ""
			optLimit = ""
			If FVatInclude = "N" Then
				optTaxCk = "N"
			Else
			 	optTaxCk = "Y"
			End If

			If Not(rsget.EOF or rsget.BOF) Then
				obj("useYn") = "Y"										'��ǰ�ɼǻ�뿩��
				obj("optType") = "1001"									'��ǰ�ɼ����� | �ܵ���(1001), ������(1002)
				Set obj("optionContentList") = jsArray()				'��ǰ�ɼǸ��
					Set obj("optionContentList")(0) = jsObject()
						obj("optionContentList")(0)("optionCdNm") = "����"
				Set obj("optionFullContentList") = jsArray()

				i = 0
				Do until rsget.EOF
				    optLimit = rsget("optLimit")
				    optLimit = optLimit-5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = 999

					Set obj("optionFullContentList")(i) = jsObject()
						obj("optionFullContentList")(i)("optionCdNm") = "����"						'��ǰ�ɼǸ�
						obj("optionFullContentList")(i)("optionContent1") = db2Html(rsget("optionname"))		'�ɼǳ���1
						obj("optionFullContentList")(i)("optionAddAmt") = optLimit 					'�ɼǼ���
						obj("optionFullContentList")(i)("optionAddBuyPrice") = getEzwelAddSuplyPrice(rsget("optaddprice"))	'�ɼǸ��԰�
						obj("optionFullContentList")(i)("optionAddPrice") =  rsget("optaddprice")	'�ɼ��߰�����
						obj("optionFullContentList")(i)("useYn") = "Y"								'�ɼǻ󼼻�뿩��
						obj("optionFullContentList")(i)("imgPath") = ""								'�ɼǽ�����̹���
						obj("optionFullContentList")(i)("imgDispYn") = "N"							'�ɼ��̹������⿩��
						obj("optionFullContentList")(i)("sortNo") = i + 1							'�ɼ����ļ���
						obj("optionFullContentList")(i)("imgDtlPath") = ""							'�ɼǻ��̹���
						obj("optionFullContentList")(i)("cspOptionFullNum") = rsget("itemoption")	'��ü�ɼǻ��ڵ�
					rsget.MoveNext
					i = i + 1
				Loop
			End If
			rsget.Close
		Else
			obj("useYn") = "N"					'��ǰ�ɼǻ�뿩��
		End If
			getEzwelItemOptionJson = obj.jsString
		Set obj = nothing
	End Function
End Class

Class CEzwel
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


	Public Sub getEzwelNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
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
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.ezwelStatCD,-9) as ezwelStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(bm.depthCode, '') as depthCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_regItem R on i.itemid=R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "									'�ö��/ȭ�����/�ؿ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
'		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (Select makerid From db_etcmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.dbo.tbl_ezwel_regItem where ezwelStatCD>3) "
		strSql = strSql & "	and uc.isExtUsing='Y'"  ''20130304 �귣�� ���޻�뿩�� Y��.
		strSql = strSql & addSql																				'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CEzwelItem
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
                FOneItem.FezwelStatCD		= rsget("ezwelStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub getEzwelEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.ezwelGoodNo, m.ezwelprice, m.ezwelSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	,isnull(bm.depthCode, '') as depthCode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' "
		strSql = strSql & "		or i.itemdiv in ('06', '16') "

		'Ȩ/���� > ��ȭ/�ö�� > �Ĺ�/�ö�� ī�װ��鼭 �ɴٹ�, �����ù� ���ϸ� ǰ��
		strSql = strSql & "		or "
		strSql = strSql & "		( "
		strSql = strSql & "			(i.cate_large = '050' and i.cate_mid = '110' and i.cate_small = '030') "
		strSql = strSql & "			AND ((i.itemname like '%�ɴٹ�%') or (i.itemname like '%�����ù�%')) "
		strSql = strSql & "		) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ezwel_regitem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ezwel_Newcate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_ezwel_Newcate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ezwelGoodNo is Not Null "									'#��� ��ǰ��
''rw strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CezwelItem
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
				FOneItem.FezwelGoodNo		= rsget("ezwelGoodNo")
				FOneItem.Fezwelprice		= rsget("ezwelprice")
				FOneItem.FezwelSellYn		= rsget("ezwelSellYn")

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

                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FdepthCode			= rsget("depthCode")
                FOneItem.Fregitemname		= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub
End Class

'Ezwel ��ǰ�ڵ� ���
Function getEzwelGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ezwelgoodno FROM db_etcmall.dbo.tbl_ezwel_regitem WHERE itemid = '"&iitemid&"' and ezwelStatcd <> '4' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getEzwelGoodno = rsget("ezwelgoodno")
	End If
	rsget.Close
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function GetEzwel10wonDown(value)
   	GetEzwel10wonDown = Fix(value/10)*10
End Function

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
	v = replace(v, "", "&gt;")
	'v = replace(v, ":", "")			'http:// �� :�� ġȯ�ǹǷ� �н�
    rpTxt = v
End Function

Function rpContent(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "<script>", "")
    v = replace(v, "</script>", "")
    v = Replace(v, "<embed>", "")
    v = Replace(v, "</embed>", "")
    v = Replace(v, "<body>", "")
    v = Replace(v, "</body>", "")
    v = replace(v, "<iframe>", "")
    v = replace(v, "</iframe>", "")
    v = replace(v, "<meta>", "")
    v = replace(v, "</meta>", "")
	v = replace(v, "<object>", "")
	v = replace(v, "</object>", "")
	v = replace(v, "<style>", "")
	v = replace(v, "</style>", "")
	v = replace(v, "<link>", "")
	v = replace(v, "</link>", "")
	v = replace(v, "<base>", "")
	v = replace(v, "</base>", "")
	v = replace(v, "<applet>", "")
	v = replace(v, "</applet>", "")
    rpContent = v
End Function

Function GetEzwelBuyPrice(value)
   	GetEzwelBuyPrice = Clng(value - (value / CEzwelMARGIN))
End Function

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function getAccessToken()
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 isnull(accessToken, '') as accessToken, lastupdate "&VbCRLF
	strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	strSql = strSql & " WHERE mallid='"& CMALLNAME &"'"&VbCRLF
	strSql = strSql & " and inikey = 'auth'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		getAccessToken	= rsget("accessToken")
	End If
	rsget.close
End Function
%>
