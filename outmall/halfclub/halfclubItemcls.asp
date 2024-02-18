<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "halfclub"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST APIURL = "http://api.tricycle.co.kr"
CONST UPCHECODE = "A5703"								'��ü�ڵ�
CONST APIKEY = "B6D75816-1F35-4450-8B9B-71137B9212F9"	'API KEY
CONST CDEFALUT_STOCK = 999

Class CHalfclubItem
	Public Fitemid
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
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Ficon2Image
	Public Fsourcearea
	Public Fmakername
	Public FBrandName
	Public FBrandNameKor
	Public FItemsize
	Public FItemsource
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FHalfClubStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FHalfclubGoodno
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public FBrandCode
	Public FNeedInfoDiv
	Public FItemweight
	Public Fcdmkey
	Public Fcddkey

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

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
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

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

	Public Function getMatchingInfoDiv(halfClubInfoDiv)
		Dim mappingDiv
		Select Case halfClubInfoDiv
			Case "C01"		mappingDiv = "01"
			Case "C02"		mappingDiv = "02"
			Case "C03"		mappingDiv = "03"
			Case "C04"		mappingDiv = "04"
			Case "C05"		mappingDiv = "05"
			Case "C06"		mappingDiv = "06"
			Case "C07"		mappingDiv = "17"
			Case "C08"		mappingDiv = "18"
			Case "C09"		mappingDiv = "19"
			Case "C10"		mappingDiv = "23"
			Case "C11"		mappingDiv = "25"
			Case "C12"		mappingDiv = "26"
			Case "C13"		mappingDiv = "08"
			Case "C14"		mappingDiv = "21"
			Case "C20"		mappingDiv = "35"
		End Select
		If FinfoDiv = mappingDiv Then
			getMatchingInfoDiv = "Y"
		Else
			getMatchingInfoDiv = "N"
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		If specialPrice <> "" Then
			MustPrice = specialPrice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If (GetTenTenMargin < CMAXMARGIN) Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
			MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
		End If
	End Function

	'�ִ� ���� ����
	Public Function getLimitHalfClubEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 1000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimitHalfClubEa = ret
	End Function

	'// ����Ŭ�� �Ǹſ��� ��ȯ
	Public Function gethalfclubSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				gethalfclubSellYn = "Y"
			Else
				gethalfclubSellYn = "N"
			End If
		Else
			gethalfclubSellYn = "N"
		End If
	End Function

	Public Function getItemidYear()
		If Clng(Fitemid) <= 1199999 Then
			getItemidYear = "2014"
		ElseIf Clng(Fitemid) >= 1200000 AND Clng(Fitemid) <= 1399999 Then
			getItemidYear = "2015"
		ElseIf Clng(Fitemid) >= 1400000 AND Clng(Fitemid) <= 1599999 Then
			getItemidYear = "2016"
		ElseIf Clng(Fitemid) >= 1600000 AND Clng(Fitemid) <= 1799999 Then
			getItemidYear = "2017"
		ElseIf Clng(Fitemid) >= 1800000 Then
			getItemidYear = Year(Date())
		End If
	End Function

    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST��ǰ] "&FItemName
		Else
			buf = "["&FBrandNameKor&"] "&FItemName
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"&","��")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
        buf = LeftB(buf, 100)
        getItemNameFormat = buf
    end function

    public function getOptionNameFormat(v)
        dim buf
        buf = replace(v,"&"," ")
        buf = replace(buf,"(","")
        buf = replace(buf,")","")
        buf = replace(buf,"/","")
        buf = replace(buf,"-","")
        buf = replace(buf,"+","_")
		buf = replace(buf,"[","")
		buf = replace(buf,"]","")
        getOptionNameFormat = buf
    end function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

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

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getHalfClubContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_halfclub.jpg'></p><br>"

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

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
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_halfclub.jpg"">")
		getHalfClubContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = ""
			strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_halfclub.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_halfclub.jpg"">")
			getHalfClubContParamToReg = strRst
		End If
		rsget.Close
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

	Public Function getHalfClubAddImageParam()
		Dim strRst, strSQL, i, k, tmpCnt, addImgUrl
		strRst = ""
		strRst = strRst & " <ImgURL_Base>"&FbasicImage&"/10x10/thumbnail/500!x500!/quality/85/"&"</ImgURL_Base>"			'#��ǰ �⺻ ū �̹��� URL(350 �̻�)
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If (IsNULL(rsget("addimage_600")) or (rsget("addimage_600")="")) Then
					addImgUrl = "add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				Else
					addImgUrl = "add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")
				End If

				If rsget("imgType") = "0" Then
					strRst = strRst & "	<ImgURL_Other"&i&">http://webimage.10x10.co.kr/image/"&addImgUrl&"/10x10/thumbnail/500!x500!/quality/85/</ImgURL_Other"&i&">"					'�߰� �̹��� 1 URL
					tmpCnt = tmpCnt + 1
				End If
				rsget.MoveNext
				If i>=3 Then Exit For
			Next
		End If
'rw tmpCnt
'response.end
		If tmpCnt < 3 Then
			For k = tmpCnt + 1 to 3
				strRst = strRst & "	<ImgURL_Other"&k&" />"
			Next
		End If
		rsget.Close
		getHalfClubAddImageParam = strRst
	End Function

	Public Function getHalfClubOptParamtoREG()
		Dim strSql, strRst, vItemOption, vOptionName, vOptAddPrice, vOptLimit, i
		strRst = ""
		vOptAddPrice		= 0
		strSql = ""
		strSql = strSql & " SELECT TOP 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(100),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, (optlimitno-optlimitsold) as optLimit " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' and o.optsellyn='Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " and o.optionname = (SELECT db_etcmall.[dbo].[RemoveSpecialChars](o.optionname)) "
		strSql = strSql & " ORDER BY o.itemoption ASC "
		rsget.Open strSql, dbget, 1
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = strRst & "			<OptionInfo>"															'�ɼ� ���� ���� ������Ʈ
			For i = 1 to rsget.RecordCount
				vItemOption	 		= rsget("itemoption")
				vOptionName 		= db2Html(rsget("optionname"))
				vOptAddPrice		= rsget("optaddprice")
				vOptLimit			= rsget("optLimit")
				vOptLimit			= vOptLimit - 5
				If (vOptLimit < 1) Then vOptLimit = 0
				If (FLimitYN <> "Y") Then vOptLimit = CDEFALUT_STOCK

				strRst = strRst & "				<Option>"
				strRst = strRst & "					<OptCd>"&vItemOption&"</OptCd>"								'#�ɼ��ڵ�
				strRst = strRst & "					<OptNm><![CDATA["&getOptionNameFormat(vOptionName)&"]]></OptNm>"								'#�ɼǸ�
				strRst = strRst & "					<OptPri>"&vOptAddPrice&"</OptPri>"							'�ɼǰ� (�ű� �Ķ����)
				strRst = strRst & "					<InvQty>"&vOptLimit&"</InvQty>"								'�ɼ� ��� ����(�Ǹ� ���� �� ���� 0)
				strRst = strRst & "				</Option>"
				rsget.MoveNext
			Next
			strRst = strRst & "			</OptionInfo>"
		Else
			strRst = strRst & "			<OptionInfo>"
			strRst = strRst & "				<Option>"
			strRst = strRst & "					<OptCd>0000</OptCd>"								'#�ɼ��ڵ�
			strRst = strRst & "					<OptNm>���ϻ�ǰ</OptNm>"								'#�ɼǸ�
			strRst = strRst & "					<OptPri>0</OptPri>"									'�ɼǰ� (�ű� �Ķ����)
			strRst = strRst & "					<InvQty>"&getLimitHalfClubEa()&"</InvQty>"			'�ɼ� ��� ����(�Ǹ� ���� �� ���� 0)
			strRst = strRst & "				</Option>"
			strRst = strRst & "			</OptionInfo>"
		End If
		rsget.Close
		getHalfClubOptParamtoREG = strRst
	End Function

	Public Function getHalfClubItemInfoCdParameter()
		Dim strRst
		Dim strSql, buf, isMatchInfoDiv
		Dim mallinfoCd, infoContent, mallinfodiv, vType

		isMatchInfoDiv = getMatchingInfoDiv(FNeedInfoDiv)
		If isMatchInfoDiv = "N" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 100 mallinfoCd, infoContent "
			strSql = strSql & " FROM db_etcmall.[dbo].[tbl_halfclub_fakeInfoCodeMap] "
			strSql = strSql & " WHERE mallinfoDiv = '"&FNeedInfoDiv&"' "
		Else
			strSql = ""
			strSql = strSql & " SELECT top 100 M.* , "
			strSql = strSql & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND isnull(IC.safetyNum, '') <> ''  THEN IC.safetyNum "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND isnull(IC.safetyNum, '') = ''  THEN tr.certNum "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' "
			strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'I') THEN '�󼼼��� ǥ��' "
			strSql = strSql & " 	 WHEN (M.infoCd='10000') THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' "
			strSql = strSql & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035' "
			strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
			strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
			strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
			strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
			strSql = strSql & " LEFT JOIN ( "
			strSql = strSql & "  SELECT TOP 1 itemid, certNum FROM db_item.dbo.tbl_safetycert_tenReg where itemid = '"&FItemID&"' "
			strSql = strSql & " ) as tr on I.itemid = tr.itemid "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
			strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
			strSql = strSql & " WHERE M.mallid = 'halfclub' and IC.itemid='"&FItemID&"' "
		End If

		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			buf = buf & " 		<NotiInfo>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
			    	infoContent = replace(infoContent, "/", " ")
			    Else
			    	infoContent = "�󼼼��� ����"
				End If
				buf = buf & "			<Noti>"
				buf = buf & "				<NotiNv><![CDATA["&mallinfoCd&"]]></NotiNv>"
				buf = buf & "				<NotiValue><![CDATA["&infoContent&"]]></NotiValue>"
				buf = buf & "			</Noti>"
				rsget.MoveNext
			Loop
			buf = buf & "		</NotiInfo>"
		End If
		rsget.Close
		getHalfClubItemInfoCdParameter = buf
		'rw buf
	End Function

	Public Function getCertInfoParam()
		Dim strRst, strSql, i, arrRows, notarrRows, newCertNo, nLp, newDiv, tCode, SafeCertTarget
		Dim buf
		strSql = ""
		strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
		strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows()
		Else
			notarrRows = "Y"
		End If
		rsget.Close

		If notarrRows = "" Then		'���ȹ� ����� �����Ͷ�� ����� �ű�
			If FsafetyYn = "Y" Then
				SafeCertTarget = "RequireCert"
				For nLp =0 To UBound(arrRows,2)
			    	newDiv = ""
					Select Case arrRows(1,nLp)
						Case "10"				'�����ǰ > ��������
							newDiv = "Electric"
							tCode = "SafeCert"
						Case "20"				'�����ǰ > ����Ȯ�� �Ű�
							newDiv = "Electric"
							tCode = "SafeCheck"
						Case "30"				'�����ǰ > ������ ���ռ� Ȯ��
							newDiv = "Electric"
							tCode = "SupplierCheck"
						Case "40"				'��Ȱ��ǰ > ��������
							newDiv = "Living"
							tCode = "SafeCert"
						Case "50"				'��Ȱ��ǰ > ��������Ȯ��
							newDiv = "Living"
							tCode = "SafeCheck"
						Case "60"				'��Ȱ��ǰ > ����ǰ��ǥ��
							newDiv = "Living"
							tCode = "SupplierCheck"
						Case "70"				'�����ǰ > ��������
							newDiv = "Child"
							tCode = "SafeCert"
						Case "80"				'�����ǰ > ����Ȯ��
							newDiv = "Child"
							tCode = "SafeCheck"
						Case "90"				'�����ǰ > ������ ���ռ� Ȯ��
							newDiv = "Child"
							tCode = "SupplierCheck"
					End Select

					newCertNo = arrRows(0,nLp)
					If newCertNo = "x" Then
						newCertNo = ""
					End If

					strRst = strRst & "	<CertInfo>"
					strRst = strRst & "		<TargetCode>"&tCode&"</TargetCode>"
					strRst = strRst & "		<SafeCertType>"&newDiv&"</SafeCertType>"
					strRst = strRst & "		<CertNum><![CDATA["&newCertNo&"]]></CertNum>"				'������ȣ
					strRst = strRst & "	</CertInfo>"
				Next
			Else
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'������ȣ
				strRst = strRst & "	</CertInfo>"
			End If
		Else
			If FsafetyYn = "Y" AND FSafetyNum <> "" Then
'				SafeCertTarget = "RequireCert"
'				Select Case FsafetyDiv
'					Case "10"			'[����ǰ] ��������
'						newDiv = "Living"
'						tCode = "SafeCert"
'					Case "20"			'[�����ǰ] ��������
'						newDiv = "Electric"
'						tCode = "SafeCert"
'					Case "30"			'[����ǰ] ����/ǰ��ǥ��
'						newDiv = "Living"
'						tCode = "SupplierCheck"
'					Case "40"			'[����ǰ] ��������Ȯ��
'						newDiv = "Living"
'						tCode = "SafeCheck"
'					Case "50"			'[����ǰ] ��̺�ȣ����
'						newDiv = "Child"
'						tCode = "ProtectedPackage"
'				End Select
'				strRst = strRst & "	<CertInfo>"
'				strRst = strRst & "		<TargetCode>"&tCode&"</TargetCode>"
'				strRst = strRst & "		<SafeCertType>"&newDiv&"</SafeCertType>"
'				strRst = strRst & "		<CertNum><![CDATA["&FSafetyNum&"]]></CertNum>"				'������ȣ
'				strRst = strRst & "	</CertInfo>"
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'������ȣ
				strRst = strRst & "	</CertInfo>"
			Else
				SafeCertTarget = "NotCert"
				strRst = strRst & "	<CertInfo>"
				strRst = strRst & "		<TargetCode>SafeCert</TargetCode>"
				strRst = strRst & "		<SafeCertType>NONE</SafeCertType>"
				strRst = strRst & "		<CertNum></CertNum>"				'������ȣ
				strRst = strRst & "	</CertInfo>"
			End If
		End If

		buf = ""
		buf = buf & " 			<SafeCertTarget>"&SafeCertTarget&"</SafeCertTarget>"
		If SafeCertTarget <> "NotCert" Then
			buf = buf & "			<CertInfos>"
			buf = buf & strRst
			buf = buf & "			</CertInfos>"
		End If
		getCertInfoParam = buf
	End Function

	Public Function getItemweight()
		Dim itemweight
		If FItemweight <> 0 Then
			itemweight = FItemweight / 1000
			If itemweight < 0.01 Then
				itemweight = 0
			End If
		Else
			itemweight = 0
		End If
		getItemweight = itemweight
	End Function

	'�⺻���� ��� XML
	Public Function getHalfClubItemRegParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
		strRst = strRst & "<soap12:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & " </SOAPHeaderAuth>"
		strRst = strRst & "</soap12:Header>"
		strRst = strRst & "<soap12:Body>"
		strRst = strRst & "	<Set_GoodsRegister xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<req_Goods>"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"												'#��ǰ �ڵ�
		strRst = strRst & "			<CategoryCd></CategoryCd>"												'ī�װ� �ڵ�
		strRst = strRst & "			<CategoryNm></CategoryNm>"												'ī�װ� ��
		strRst = strRst & "			<BrdCd>"&FBrandCode&"</BrdCd>"											'#�귣�� �ڵ�(����Ŭ�� ����)
		strRst = strRst & "			<BrdNm><![CDATA[�ٹ�����]]></BrdNm>"										'�귣�� ��
		strRst = strRst & "			<Item_BCode></Item_BCode>"												'����Ŭ�� ���� ��з� �ڵ�
		strRst = strRst & "			<Item_BName></Item_BName>"												'����Ŭ�� ���� ��з� ��
		strRst = strRst & "			<Item_MCode></Item_MCode>"												'����Ŭ�� ���� �ߺз� �ڵ�
		strRst = strRst & "			<Item_MName></Item_MName>"												'����Ŭ�� ���� �ߺз� ��
		strRst = strRst & "			<Item_SCode>"&FDepthCode&"</Item_SCode>"								'#����Ŭ�� ���� �Һз� �ڵ�
		strRst = strRst & "			<Item_SName></Item_SName>"												'����Ŭ�� ���� �Һз� ��
		strRst = strRst & "			<MakeYear>"&getItemidYear()&"</MakeYear>"								'#��ǰ �����⵵
		strRst = strRst & "			<PrdNm><![CDATA["&getItemNameFormat()&"]]></PrdNm>"						'#��ǰ��
		strRst = strRst & "			<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#��ǰ ����(�ð�) / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "			<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#��ǰ �ǸŰ� / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "			<PrdDescInfo><![CDATA["&getHalfClubContParamToReg()&"]]></PrdDescInfo>"	'#��ǰ �� ����
		strRst = strRst & "			<CopyInfo></CopyInfo>"													'��ǰ ī�Ǹ�
		strRst = strRst & "			<Nation><![CDATA["&Fsourcearea&"]]></Nation>"							'#��ǰ ������
		strRst = strRst & getHalfClubAddImageParam()
		strRst = strRst & "			<PrdWeight>"&getItemweight()&"</PrdWeight>"								'��ǰ ����(���� : kg)
		strRst = strRst & "			<SalOut>a</SalOut>"														'#��ǰ ����(�Ǹ��� : a, �Ͻ�ǰ�� : b, �Ǹ����� : k)
		strRst = strRst & "			<ImageUpdate>Y</ImageUpdate>"											'�̹��� ��� ����(Y : ���, N : �̵��)
		strRst = strRst & getHalfClubOptParamtoREG()
		strRst = strRst & getHalfClubItemInfoCdParameter()
		strRst = strRst & getCertInfoParam()
		strRst = strRst & "			<IsConversion></IsConversion>"											'ȯ�ݼ� ��ǰ ���� (1 : ȯ�ݼ���ǰ, 0 : ȯ�ݼ� ��ǰ �ƴ�)
		strRst = strRst & "		</req_Goods>"
		strRst = strRst & "	</Set_GoodsRegister>"
		strRst = strRst & "</soap12:Body>"
		strRst = strRst & "</soap12:Envelope>"
		getHalfClubItemRegParameter = strRst
'response.write replace(strRst, "UTF-8","EUC-KR")
'response.write replace(strRst, "?xml","aaaass")
'response.end
	End Function

	'��ǰ���� ���� XML
	Public Function getHalfClubItemEditParameter(ichgSellyn)
		Dim strRst, SalOut
		Select Case ichgSellyn
			Case "Y"	SalOut = "a"
			'Case "N"	SalOut = "b"	'�Ͻ�ǰ��
			Case "N"	SalOut = "k"	'�Ǹ�����..2019-03-18 16:29 ������..b�� ó���� ���� �߻� x �׷��� ���� ǰ���� ���� �ʴ� Case �߻��Ͽ� k�� ����
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
		strRst = strRst & "<soap12:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & " </SOAPHeaderAuth>"
		strRst = strRst & "</soap12:Header>"
		strRst = strRst & "<soap12:Body>"
		strRst = strRst & "	<Set_GoodsRegister xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<req_Goods>"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"												'#��ǰ �ڵ�
		strRst = strRst & "			<CategoryCd></CategoryCd>"												'ī�װ� �ڵ�
		strRst = strRst & "			<CategoryNm></CategoryNm>"												'ī�װ� ��
		strRst = strRst & "			<BrdCd>"&FBrandCode&"</BrdCd>"											'#�귣�� �ڵ�(����Ŭ�� ����)
		strRst = strRst & "			<BrdNm><![CDATA[�ٹ�����]]></BrdNm>"										'�귣�� ��
		strRst = strRst & "			<Item_BCode></Item_BCode>"												'����Ŭ�� ���� ��з� �ڵ�
		strRst = strRst & "			<Item_BName></Item_BName>"												'����Ŭ�� ���� ��з� ��
		strRst = strRst & "			<Item_MCode></Item_MCode>"												'����Ŭ�� ���� �ߺз� �ڵ�
		strRst = strRst & "			<Item_MName></Item_MName>"												'����Ŭ�� ���� �ߺз� ��
		strRst = strRst & "			<Item_SCode>"&FDepthCode&"</Item_SCode>"								'#����Ŭ�� ���� �Һз� �ڵ�
		strRst = strRst & "			<Item_SName></Item_SName>"												'����Ŭ�� ���� �Һз� ��
		strRst = strRst & "			<MakeYear>"&getItemidYear()&"</MakeYear>"								'#��ǰ �����⵵
		strRst = strRst & "			<PrdNm><![CDATA["&getItemNameFormat()&"]]></PrdNm>"						'#��ǰ��
		strRst = strRst & "			<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#��ǰ ����(�ð�) / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "			<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#��ǰ �ǸŰ� / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "			<PrdDescInfo><![CDATA["&getHalfClubContParamToReg()&"]]></PrdDescInfo>"	'#��ǰ �� ����
		strRst = strRst & "			<CopyInfo></CopyInfo>"													'��ǰ ī�Ǹ�
		strRst = strRst & "			<Nation><![CDATA["&Fsourcearea&"]]></Nation>"							'#��ǰ ������
		strRst = strRst & "			<PrdWeight>"&getItemweight()&"</PrdWeight>"								'��ǰ ����(���� : kg)
		strRst = strRst & "			<SalOut>"&SalOut&"</SalOut>"											'#��ǰ ����(�Ǹ��� : a, �Ͻ�ǰ�� : b, �Ǹ����� : k)
		strRst = strRst & getHalfClubAddImageParam()
		strRst = strRst & "			<ImageUpdate>"&Chkiif(isImageChanged, "Y", "N")&"</ImageUpdate>"	'�̹��� ��� ����(Y : ���, N : �̵��)
		If ichgSellyn = "N" Then
			strRst = strRst & "			<OptionInfo />"
		Else
			strRst = strRst & getHalfClubOptParamtoREG()
		End If
		strRst = strRst & getHalfClubItemInfoCdParameter()
		strRst = strRst & getCertInfoParam()
		strRst = strRst & "		</req_Goods>"
		strRst = strRst & "	</Set_GoodsRegister>"
		strRst = strRst & "</soap12:Body>"
		strRst = strRst & "</soap12:Envelope>"
		getHalfClubItemEditParameter = strRst
		' if session("ssBctID")="icommang" or session("ssBctID")="kjy8517" then
		' 	response.write replace(strRst, "UTF-8","EUC-KR")
		' 	response.write replace(strRst, "?xml","aaaass")
		' End If
'response.end
	End Function

	'��ǰ���� ���� XML
	Public Function getHalfClubPriceParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "<soap:Header>"
		strRst = strRst & "	<SOAPHeaderAuth xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & " 	<User_ID>"&UPCHECODE&"</User_ID>"
		strRst = strRst & " 	<User_PWD>"&APIKEY&"</User_PWD>"
		strRst = strRst & "	</SOAPHeaderAuth>"
		strRst = strRst & "</soap:Header>"
		strRst = strRst & "<soap:Body>"
		strRst = strRst & "	<Set_Good_Price_Change xmlns=""http://api.tricycle.co.kr/"">"
		strRst = strRst & "		<gar ResultCode="""" ResultMsg="""">"
		strRst = strRst & "			<PCode>"&FItemid&"</PCode>"
		strRst = strRst & "			<goodpriinfo>"
		strRst = strRst & "				<PCode>"&FItemid&"</PCode>"
		strRst = strRst & "				<Margin>13</Margin>"
		strRst = strRst & "				<Pri>"&Clng(GetRaiseValue(ForgPrice/10)*10)&"</Pri>"					'#��ǰ ����(�ð�) / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "				<SalPri>"&Clng(GetRaiseValue(MustPrice()/10)*10)&"</SalPri>"			'#��ǰ �ǸŰ� / ¾���� �Է¾ȵ� ����(2018-10-25 ����)
		strRst = strRst & "			</goodpriinfo>"
		strRst = strRst & "		</gar>"
		strRst = strRst & "	</Set_Good_Price_Change>"
		strRst = strRst & "</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getHalfClubPriceParameter = strRst
'response.write replace(strRst, "UTF-8","EUC-KR")
'response.write replace(strRst, "?xml","aaaass")
'response.end
	End Function
End Class

Class CHalfclub
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
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'// �̵�� ��ǰ ���(��Ͽ�)
	Public Sub getHalfClubNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_item.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv, c.infodiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, isNULL(R.HalfClubStatCD,-9) as HalfClubStatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, am.brandCode, am.needInfoDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.[tbl_halfclub_cate_mapping] "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.[tbl_halfclub_cate_mapping] as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_halfclub_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '06', '08') "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'�ö��/ȭ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_halfclub_regItem WHERE HalfClubStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.	'11st��ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
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
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FHalfClubStatCD		= rsget("HalfClubStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FBrandCode			= rsget("brandCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FInfodiv 			= rsget("infodiv")
				FOneItem.FNeedInfoDiv 		= rsget("needInfoDiv")
				FOneItem.FItemweight 		= rsget("itemweight")
		End If
		rsget.Close
	End Sub

	Public Sub gethalfclubEditPriceOneItem()
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
		End If
		rsget.Close
	End Sub

	Public Sub gethalfclubEditOneItem()
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv, c.infodiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	,isNULL(m.HalfClubStatCD,-9) as HalfClubStatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, am.brandCode, am.needInfoDiv "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '06', '08', '09') "
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_halfclub_cate_mapping "
		strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.[tbl_halfclub_cate_mapping] as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_halfclub_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.HalfClubGoodNo is Not Null "		'��� ��ǰ��
		strSql = strSql & " and m.HalfclubStatCD = '7' "				'���οϷ�� �ֵ鸸 ������ �ȴ���..TEST �غ��� ��
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHalfClubItem
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
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FHalfClubStatCD		= rsget("HalfClubStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FBrandCode			= rsget("brandCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FInfodiv 			= rsget("infodiv")
				FOneItem.FNeedInfoDiv 		= rsget("needInfoDiv")
				FOneItem.FItemweight 		= rsget("itemweight")
				FOneItem.FMaySoldOut		= rsget("maySoldOut")
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

Public Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

function replaceRst(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, """", "&quot;")
	'v = replace(v, "&", "&amp;")

	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function

function replaceMsg(v)
	if IsNull(v) then
		replaceMsg = ""
		Exit function
	end if
	v = Replace(v, vbcrlf,"")
	v = Replace(v, vbCr,"")
	v = Replace(v, vbLf,"")
    replaceMsg = v
end function
%>
