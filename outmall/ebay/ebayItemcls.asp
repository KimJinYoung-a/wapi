<!-- #include virtual="/outmall/ebay/inc_gubunChk.asp"-->
<%
CONST CMAXMARGIN = 15
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST APIURL = "http://api.11st.co.kr/rest"
CONST APISSLURL = "https://sa.esmplus.com"
CONST APIkey = "a2319e071dbc304243ee60abd07e9664"
CONST CDEFALUT_STOCK = 99999

Class CEbayItem
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
	Public FAuctionStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FCateCode
	Public FSDCategoryCode
	Public Fcdmkey
	Public Fcddkey
	Public FSt11GoodNo
	Public FSt11price
	Public FSt11SellYn
	Public FIsbn13

	Public FSafeDiv
	Public FIsNeed
	Public FDepth1Code
	Public FAdultType

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

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

	'�ִ� ���� ����
	Public Function getLimitEbayEa()
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
		getLimitEbayEa = ret
	End Function

	'// 11st �Ǹſ��� ��ȯ
	Public Function getEbaySellyn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getEbaySellyn = "Y"
			Else
				getEbaySellyn = "N"
			End If
		Else
			getEbaySellyn = "N"
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

    public function getItemNameFormat(v)
        dim buf
		If application("Svr_Info") = "Dev" Then
			FItemName = "[TEST��ǰ] "&FItemName
		End If
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"&","��")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
		If v = 1 Then
	        buf = LeftB(buf, 50)
		End If
        getItemNameFormat = buf
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
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If (cntType <> cntOpt) OR (cntOpt > 2) Then		'3�� �ɼ� ��������
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

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getEbayContParamToReg(obj, vGubun)
		Dim strRst, strSQL, tmpContent, gubunStr
		If vGubun = "A" Then
			gubunStr = "auction"
		Else
			gubunStr = "gmarket"
		End If

		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_"&gubunStr&".jpg></p><br />")
		strRst = strRst & ("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & ("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & ("<tr>")
		strRst = strRst & ("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""��ǰ��"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & ("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '����', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & ("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat(2)
		strRst = strRst & ("</p>")
		strRst = strRst & ("</td>")
		strRst = strRst & ("</tr>")
		strRst = strRst & ("</table>")
		strRst = strRst & ("</div>")

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br />")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br />")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br />")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br />")

		'#��� ���ǻ���
		strRst = strRst & ("<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_"&gubunStr&".jpg>")
		strRst = strRst & ("</div>")
		tmpContent = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = rsget("textVal")
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_"&gubunStr&".jpg></p><br />" & strRst & "<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_"&gubunStr&".jpg></div>"
			tmpContent = strRst
		End If
		rsget.Close

		Set obj("itemAddtionalInfo")("descriptions") = jsObject()
			Set obj("itemAddtionalInfo")("descriptions")("kor") = jsObject()
				obj("itemAddtionalInfo")("descriptions")("kor")("type") = 2				'#��ǰ������Ÿ�� | 1 contentID(��������), 2 html
				obj("itemAddtionalInfo")("descriptions")("kor")("contentId") = ""		'��ǰ������ �ڵ� | ��ǰ������Ÿ���� 1�� �� �ʼ�
				obj("itemAddtionalInfo")("descriptions")("kor")("html")	= tmpContent	'#��ǰ������ html | iframe, Script �Ұ�
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
		If trim(Fkeywords) = "" Then Exit Function
		Fkeywords  = replace(Fkeywords,"%", "")
		Fkeywords  = replace(Fkeywords,"/", ",")
		Fkeywords  = replace(Fkeywords,chr(13), "")
		Fkeywords  = replace(Fkeywords,chr(10), "")
		Fkeywords  = replace(Fkeywords,chr(9), "")
		Fkeywords  = replace(Fkeywords,chr(32), "")

		arrRst = Split(Fkeywords,",")
		If Ubound(arrRst) = 0 then
			arrRst2 = split(arrRst(0),";")
			If Ubound(arrRst2) > 0 then
				arrRst = split(Fkeywords,";")
			End If
		End If

		If Ubound(arrRst)+1 >= 5 then
			getItemKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			getItemKeyword = keyword1
		End If
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

	Public Function getEbayImageParameter(obj)
		Dim strRst, strSQL, i, imgAdds, spImage

		Set obj("itemAddtionalInfo")("images") = jsObject()
			obj("itemAddtionalInfo")("images")("basicImgURL") = FbasicImage&"/10x10/thumbnail/600x600/quality/85/"			'#��ǰ �⺻�̹��� | �ּ� 600x600 ���� 1000x1000

 		strSQL = ""
		strSQL = strSQL & " SELECT TOP 2 gubun,ImgType,addimage_400,addimage_600,addimage_1000 "
		strSQL = strSQL & " FROM db_item.[dbo].tbl_item_addimage "
		strSQL = strSQL & " WHERE IMGTYPE = 0 "
		strSQL = strSQL & " AND itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			imgAdds = ""
			For i=1 to rsget.RecordCount
				imgAdds = imgAdds & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & ","
				rsget.MoveNext
			Next
			If Right(imgAdds,1) = "," Then
				imgAdds = Left(imgAdds, Len(imgAdds) - 1)
			End If
			spImage = Split(imgAdds, ",")

			If isArray(spImage) Then
				If Ubound(spImage) >= 0 Then
					obj("itemAddtionalInfo")("images")("addtionalImg1URL") = spImage(0)&"/10x10/thumbnail/600x600/quality/85/"
					If Ubound(spImage) = 1 Then
						obj("itemAddtionalInfo")("images")("addtionalImg2URL") = spImage(1)&"/10x10/thumbnail/600x600/quality/85/"
					Else
						obj("itemAddtionalInfo")("images")("addtionalImg2URL") = null
					End If
				End If
			End If
		Else
			obj("itemAddtionalInfo")("images")("addtionalImg1URL") = null
			obj("itemAddtionalInfo")("images")("addtionalImg2URL") = null
		End If
		rsget.Close
	End Function

	Public Function fnCertCodes(iGubun, icertNo, icertDiv, itype)
		Dim strSql, addSql, tmpVal
		If iGubun = "ELEC" Then
			addSql = addSql & " and r.safetyDiv in ('10', '20', '30') "
		ElseIf iGubun = "LIFE" Then
			addSql = addSql & " and r.safetyDiv in ('40', '50', '60') "
		Else
			addSql = addSql & " and r.safetyDiv in ('70', '80', '90') "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP 1 r.certNum "
		strSql = strSql & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 0 "
		strSql = strSql & "		  When r.safetyDiv in ('20', '50', '80') THEN 1 "
		strSql = strSql & " 	  When r.safetyDiv in ('30', '60', '90') THEN 2 end as safetyStr "
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSql = strSql & " WHERE r.itemid='"&FItemid&"' "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			icertNo		= rsget("certNum")
			icertDiv	= rsget("safetyStr")
			tmpVal		= "Y"
		Else
			icertNo		= ""
			icertDiv	= ""
			tmpVal		= "N"
		End If
		rsget.Close

		If tmpVal = "Y" Then
			If icertDiv = 2 Then
				itype = 2
			Else
				itype = 0
			End If
		Else
			itype = 1
		End If
	End Function

	Public Function getEbayCertInfoParameter(obj)
		Dim certNo, certDiv, vType
		Set obj("itemAddtionalInfo")("certInfo") = jsObject()
			obj("itemAddtionalInfo")("certInfo")("gmkt") = null										'(G���Ͽ�) ���������ڵ�
			obj("itemAddtionalInfo")("certInfo")("iac") = null										'(���ǿ�) ���������ڵ� - �Ƿ���, �����ű��, ��ǰ����������, �ǰ���ɽ�ǰ, ģȯ������ ��
			Set obj("itemAddtionalInfo")("certInfo")("safetyCerts") = jsObject()
				Call fnCertCodes("CHILD", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("type") = vType										'#����������� ��ǰ �ƴҰ�� "�������ƴ�"���� �Է� | 0 �������, 1 �������ƴ�, 2 ��ǰ�󼼺���ǥ��
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")= null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null)("certId") = certNo			'���վ�� �����ڵ�
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("child")("details")(null)("certTargetCode") = certDiv	'���վ������ǰ�� | 0 ��������, 1 ����Ȯ��, 3 ���������ռ�Ȯ��
				End If

				Call fnCertCodes("ELEC", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("type") = vType										'#����������� ��ǰ �ƴҰ�� "�������ƴ�"���� �Է� | 0 �������, 1 �������ƴ�, 2 ��ǰ�󼼺���ǥ��
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("mandatorySafetySign") = "UnknownOrNone"			'������Կ��� | BuyingAgent : ���Ŵ���, ParallelImport �������, UnknownOrNone : �ش���׾���
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details") = null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null)("certId") = certNo			'�������� �����ڵ�
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("electric")("details")(null)("certTargetCode") = certDiv'������������ǰ�� | 0 ��������, 1 ����Ȯ��, 3 ���������ռ�Ȯ��
				End If

				Call fnCertCodes("LIFE", certNo, certDiv, vType)
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("type") = vType											'#����������� ��ǰ �ƴҰ�� "�������ƴ�"���� �Է� | 0 �������, 1 �������ƴ�, 2 ��ǰ�󼼺���ǥ��
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("mandatorySafetySign") = "UnknownOrNone"				'������Կ��� | BuyingAgent : ���Ŵ���, ParallelImport �������, UnknownOrNone : �ش���׾���
				If vType = 1 Then
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details") = null
				Else
					Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details") = jsArray()
						Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null) = jsObject()
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null)("certId") = certNo				'���ջ�Ȱ��ǰ �����ڵ�
							obj("itemAddtionalInfo")("certInfo")("safetyCerts")("life")("details")(null)("certTargetCode") = certDiv	'���ջ�Ȱ��ǰ����ǰ�� | 0 ��������, 1 ����Ȯ��, 3 ���������ռ�Ȯ��
				End If
				Set obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful") = jsObject()
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful")("type") = 2											'�������ط����ǰ����Ÿ�� | ���������� ����Ʈ�� �󼼼���ǥ��� ����, 0 �������, 1 �������ƴ�, 2 ��ǰ�󼼺���ǥ��
					obj("itemAddtionalInfo")("certInfo")("safetyCerts")("harmful")("certId") = null										'�����ڰ��˻��ȣ | type > 0�� �� �ʼ�
	End Function

	Public Function get11stOptionParam()
		Dim strSql, strRst1, strRst2, strRst3, i, optNm, optDc, chkMultiOpt, optLimit, arrList1, arrList2, optCode, j, optionname, optaddprice, tmpoName
		Dim voptsellyn, visusing
		strRst1 = ""
		strRst2 = ""
		strRst3 = ""
		chkMultiOpt = false
		If FoptionCnt > 0 Then
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql,dbget,1
		    if not rsget.Eof then
		    	chkMultiOpt = true
		        arrList1 = rsget.getRows()
		    end if
		    rsget.close
		End If
		strRst1 = strRst1 & "	<optSelectYn>Y</optSelectYn>"									'������ �ɼ� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
		strRst1 = strRst1 & "	<txtColCnt>1</txtColCnt>"										'������ | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ɼ��� ����Ͻ� ��� 1 �������� �ּž� �մϴ�.
'		If chkMultiOpt = True Then
'		strRst1 = strRst1 & "	<optionAllQty>"&getLimit11stEa&"</optionAllQty>"				'��Ƽ�ɼ� �ϰ������� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� ��� ���� ������ api ������ �Ұ��մϴ�. �ϰ������� ����.
'		strRst1 = strRst1 & "	<optionAllAddPrc>0</optionAllAddPrc>"							'��Ƽ�ɼ� �ɼǰ� 0�� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� �ɼǰ� ������ api ������ �Ұ��մϴ�. 0���� �Է� ����
'		strRst1 = strRst1 & "	<optionAllAddWght/>"											'��Ƽ�ɼ� �ϰ��ɼ��߰����� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� �ɼǹ��� ������ api ������ �Ұ��մϴ�. �ϰ������� ����.
'		End If
		strRst1 = strRst1 & "	<prdExposeClfCd>00</prdExposeClfCd>"							'��ǰ�� �ɼǰ� ���� ��� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 00 : ��ϼ�, 01 : �ɼǰ� �����ټ�, 02 : �ɼǰ� ������ ����, 03 : �ɼǰ��� ���� ��, 04 : �ɼǰ��� ���� ��
		strRst1 = strRst1 & "	<optUpdateYn>Y</optUpdateYn>"
'		strRst1 = strRst1 & "	<optMixYn/>"													'��ü�ɼ� ���տ��� |"�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������ ��ü �ɼǰ��� ���յǾ� ��Ƽ�ɼ����� ���, N : �ɼ� ����Key�� �����ϴ�(���� ��) �����θ� ��Ƽ �ɼ� ���
		If chkMultiOpt = True Then
			strSql = "select typeseq, optionTypeName, optionKindName, optaddPrice from db_item.[dbo].[tbl_item_option_Multiple] where itemid = " & FItemid & " ORDER BY Typeseq, kindSeq "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				arrList2 = rsget.getRows()
			End If
			rsget.close

			For i = 0 To Ubound(arrList1, 2)
				strRst2 = strRst2 & "	<ProductRootOption>"											'ProductRootOption |"�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
				strRst2 = strRst2 & "		<colTitle>"&arrList1(2,i)&"</colTitle>"						'�ɼǸ� | 40Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�.
				For j = 0 to Ubound(arrList2, 2)
					If arrList1(1,i) = arrList2(0,j) then

						arrList2(2,j) = replace(arrList2(2,j), "&", "+")			'2017-06-05 ������..���� �븮 ��û���� &->+�� ����

						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<colOptPrice>0</colOptPrice>"					'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "			<colValue0>"&arrList2(2,j)&"</colValue0>"		'�ɼǰ� | 50Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�. �� ��ǰ�ȿ��� �ɼǰ��� �ߺ��� �ɼ� �����ϴ�.
						strRst2 = strRst2 & "		</ProductOption>"
					end if
				Next
				strRst2 = strRst2 & "	</ProductRootOption>"
			Next

			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, isUsing, optsellyn "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
'			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "		'2017-05-17 ������ ����
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				strRst2 = strRst2 & "	<ProductOptionExt>"												'ProductOptionExt | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    tmpoName	= db2html(rsget("optionname"))
				    visUsing= rsget("isUsing")
				    voptsellyn= rsget("optsellyn")

					optionname = ""
					For i = 0 To Ubound(arrList1, 2)
						If Ubound(Split(tmpoName, ",")) > 0 Then				'2017-06-15 ������ �߰�
							optionname = optionname & arrList1(2,i) &":"&Split(tmpoName, ",")(i) &"��"
						End If
					Next

					If Right(optionname,1) = "��" Then
						optionname = Left(optionname, Len(optionname) - 1)
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
			    	If voptsellyn <> "Y" Then optLimit = 0	'2017-05-17 ������ ����
			    	If visUsing <> "Y" Then optLimit = 0	'2017-05-17 ������ ����

			'		If optLimit > 0 Then		'2017-05-17 ������ ����
						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<useYn>"&chkiif(optLimit>0,"Y","N")&"</useYn>"
						strRst2 = strRst2 & "			<colOptPrice>"&optaddprice&"</colOptPrice>"				'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "			<colCount>"&optLimit&"</colCount>"			'�ɼ������� | ��Ƽ�ɼ��� ���� �ϰ������� �ǹǷ� �Է��Ͻø� �ȵ˴ϴ�. �ɼǻ���(useYn)�� N�� ���� 0 �Է� �����մϴ�.
						strRst2 = strRst2 & "			<colSellerStockCd>"&optCode&"</colSellerStockCd>"		'��������ȣ | ������ ����ϴ� ����ȣ
						strRst2 = strRst2 & "			<optionMappingKey>"&optionname&"</optionMappingKey>"	'�ɼǸ���Key | ��Ƽ�ɼ��� ���յ� �ɼ��� �����ϱ� ���� Key(��: �ɼǸ�1:�ɼǰ�1�ӿɼǸ�2:�ɼǰ�2)
						strRst2 = strRst2 & "		</ProductOption>"
			'		End If
					rsget.MoveNext
				Loop
				strRst2 = strRst2 & "	</ProductOptionExt>"
			End If
			rsget.Close
		Else
			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				If db2Html(rsget("optionTypeName"))<>"" Then
					optNm = Replace(db2Html(rsget("optionTypeName")),":","")
				Else
					optNm = "�ɼ�"
				End If
				strRst1 = strRst1 & "	<colTitle>"&optNm&"</colTitle>"
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    optionname	= db2html(rsget("optionname"))
				    optionname = replace(optionname, "&", "+")			'2017-06-05 ������..���� �븮 ��û���� &->+�� ����
				    optionname = replace(optionname, ",", "+")
				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					If optLimit > 0 Then
						strRst2 = strRst2 & "	<ProductOption>"								'ProductOption | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
						strRst2 = strRst2 & "		<useYn>Y</useYn>"							'�ɼǻ��� | ��Ƽ�ɼ��� ���� �������� �ʴ� ����Դϴ�. Y : �����, N : ǰ��
						strRst2 = strRst2 & "		<colOptPrice>"&optaddprice&"</colOptPrice>"	'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "		<colValue0>"&optionname&"</colValue0>"		'�ɼǰ� | 50Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�. �� ��ǰ�ȿ��� �ɼǰ��� �ߺ��� �ɼ� �����ϴ�
						strRst2 = strRst2 & "		<colCount>"&optLimit&"</colCount>"			'�ɼ������� | ��Ƽ�ɼ��� ���� �ϰ������� �ǹǷ� �Է��Ͻø� �ȵ˴ϴ�. �ɼǻ���(useYn)�� N�� ���� 0 �Է� �����մϴ�.
						strRst2 = strRst2 & "		<colSellerStockCd>"&optCode&"</colSellerStockCd>"	'��������ȣ | ������ ����ϴ� ����ȣ
						strRst2 = strRst2 & "	</ProductOption>"
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If

		If FitemDiv = "06" Then
			strRst3 = strRst3 & "	<ProductCustOption>"										'�ɼǵ�� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �������ۼ��� �ɼ��� ��� �ִ� 5������ ��� ����
			strRst3 = strRst3 & "		<colOptName>�ؽ�Ʈ�� �Է��ϼ���</colOptName>"				'������ �ۼ��� �ɼǸ� | �ɼǸ� �ִ� �������� �ѱ�10��/����(����)20��)���� �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�.
			strRst3 = strRst3 & "		<colOptUseYn>Y</colOptUseYn>"							'�ɼ� ��� ���� | Y : �����, N : ������
			strRst3 = strRst3 & "	</ProductCustOption>"
		End If
		get11stOptionParam = strRst1 & strRst2 & strRst3
'rw get11stOptionParam
'response.end
	End Function

	Public Function getEbayInfoCdParameter(obj)
		Dim strSQL
		Dim mallinfodiv, mallinfoCd, infoContent, certNum
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 isNull(r.certNum, '') as certNum "
		strSQL = strSQL & "	,Case When r.safetyDiv in ('10', '40', '70') THEN 'SafeCert' "
		strSQL = strSQL & "		  When r.safetyDiv in ('20', '50', '80') THEN 'SafeCheck' "
		strSQL = strSQL & " 	  When r.safetyDiv in ('30', '60', '90') THEN 'SupplierCheck' end as safetyStr "
		strSQL = strSQL & " ,convert(date, f.certDate) as certDate, f.modelName " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.[dbo].[tbl_safetycert_info] as f on r.itemid = f.itemid " & vbcrlf
		strSQL = strSQL & " WHERE r.itemid='"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			certNum		= rsget("certNum")
		End If
		rsget.Close

		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		If certNum = "" Then
			strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN  IC.safetyNum " & vbcrlf
		Else
			If certNum = "x" Then
				certNum = "�ش����"
			End If
			strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN '"& certNum &"' " & vbcrlf
		End If
		strSql = strSql & "		 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '������ ����ǥ��' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035'  " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"'  " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='" & FItemid &"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'ebay' and IC.itemid='"&FItemid&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = CInt(rsget("mallinfodiv"))
			Set obj("itemAddtionalInfo")("officialNotice") = jsObject()
				obj("itemAddtionalInfo")("officialNotice")("officialNoticeNo") = mallinfodiv	'#��ǰ������� ��ǰ���ڵ�
				Set obj("itemAddtionalInfo")("officialNotice")("details") = jsArray()
			Do until rsget.EOF
				mallinfoCd  = rsget("mallinfoCd")
				infoContent = rsget("infoContent")
					Set obj("itemAddtionalInfo")("officialNotice")("details")(null) = jsObject()
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("officialNoticeItemelementCode") = mallinfoCd	'#��ǰ������� �׸��ڵ�
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("value") = infoContent							'#��ǰ������� ��
						obj("itemAddtionalInfo")("officialNotice")("details")(null)("isExtraMark") = false							'��ǰ������� �߰��Է¿��� | true : �߰��Է°��, false : �߰� �Է¾���
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Function

	'�⺻���� ��� XML
	Public Function getEbayItemRegParameter(vGubun)
		Dim strRst
		Dim obj
		Set obj = jsObject()
'			Set obj("isSell") = jsObject()
'				obj("isSell")(Chkiif(vGubun="A", "iac", "gmkt")) = false						'#�ǸŻ��º��� | true �Ǹ�, false �Ǹ�����, �Ǹ������� 1���� ������ ������

			Set obj("itemBasicInfo") = jsObject()
				Set obj("itemBasicInfo")("goodsName") = jsObject()
					obj("itemBasicInfo")("goodsName")("eng") = null								'������ǰ��
					obj("itemBasicInfo")("goodsName")("chi") = null								'�߹���ǰ��
					obj("itemBasicInfo")("goodsName")("jpn") = null								'�Ϲ���ǰ��
					obj("itemBasicInfo")("goodsName")("kor") = ""&getItemNameFormat(1)&""			'#�˻��� ������ǰ��
					obj("itemBasicInfo")("goodsName")("promotion") = null						'���θ�ǿ� ������ǰ��
				Set obj("itemBasicInfo")("category") = jsObject()
					Set obj("itemBasicInfo")("category")("site") = jsArray()
						Set obj("itemBasicInfo")("category")("site")(null) = jsObject()
							obj("itemBasicInfo")("category")("site")(null)("siteType") = Chkiif(vGubun="A", "1", "2")	'#G����/���� ī�װ� ����� ���� ����Ʈ ���� | 1 ����, 2 G����
							obj("itemBasicInfo")("category")("site")(null)("catCode") = ""& FCateCode &""				'#G����/���ǿ��� �����ϴ� ������(Leaf)ī�װ� �ڵ� ���
					Set obj("itemBasicInfo")("category")("shop") = jsArray()
						obj("itemBasicInfo")("category")("shop") = null
'					Set obj("itemBasicInfo")("category")("shop") = jsArray()
'						Set obj("itemBasicInfo")("category")("shop")(null) = jsObject()
'							obj("itemBasicInfo")("category")("shop")(null)("siteType") = ""			'�̴ϼ� ī�װ� ����Ʈ ����
'							obj("itemBasicInfo")("category")("shop")(null)("largeCatCode") = ""		'�̴ϼ� ��ī�װ��ڵ�
'							obj("itemBasicInfo")("category")("shop")(null)("middleCatCode") = ""	'�̴ϼ� ��ī�װ��ڵ�
'							obj("itemBasicInfo")("category")("shop")(null)("smallCatCode") = ""		'�̴ϼ� ��ī�װ��ڵ�
					Set obj("itemBasicInfo")("category")("esm") = jsObject()
						obj("itemBasicInfo")("category")("esm")("catCode") = ""& FSDCategoryCode &""	'#ESMī�װ��ڵ���

 				If FIsbn13 <> "" Then
					Set obj("itemBasicInfo")("book") = jsObject()
						obj("itemBasicInfo")("book")("isUseIsbnCode") = true						'(������ǰ��)ISBN�ڵ� ��뿩��
						obj("itemBasicInfo")("book")("isbnCode") = ""&FIsbn13&""					'(������ǰ��)ISBN�ڵ�
						obj("itemBasicInfo")("book")("price") = null								'(������ǰ��)������
						obj("itemBasicInfo")("book")("attributeCode") = null						'(������ǰ��/G��������)�߰���� ī�װ�
				End If
					Set obj("itemBasicInfo")("catalog") = jsObject()
						obj("itemBasicInfo")("catalog")("modelName") = null							'�𵨸�
						obj("itemBasicInfo")("catalog")("brandNo") = 0								'�귣���ڵ�
						obj("itemBasicInfo")("catalog")("barCode") = null							'���ڵ�
						Set obj("itemBasicInfo")("catalog")("epinCode") = jsArray()
							obj("itemBasicInfo")("catalog")("epinCode")(null) = 0					'ESM ��ǰ�з��ڵ� | ���� API �������� �ʾ� null�� ȣ��

			Set obj("itemAddtionalInfo") = jsObject()
				Set obj("itemAddtionalInfo")("buyableQuantity") = jsObject()
					obj("itemAddtionalInfo")("buyableQuantity")("type") = 0							'#���ż������� Ÿ�� | 0 : ���ż������Ѿ���, 1 : 1ȸ�� �ִ� ���ż���, 2 : ID�� �ִ� ���ż���, 3 : �Ⱓ�� �ִ� ���ż���
					obj("itemAddtionalInfo")("buyableQuantity")("qty") = null						'�ִ뱸�ż��� | ���ż������� Ÿ���� 1~3�� �� �ʼ�
					obj("itemAddtionalInfo")("buyableQuantity")("unitDate") = null					'���ѱⰣ | ���ż������� Ÿ���� 3�� �� �ʼ�
				Set obj("itemAddtionalInfo")("price") = jsObject()									'#����/G���� �ǸŰ��� | 10�������� ���
					obj("itemAddtionalInfo")("price")(Chkiif(vGubun="A", "Iac", "Gmkt")) = Clng(GetRaiseValue(MustPrice/10)*10)
				Set obj("itemAddtionalInfo")("stock") = jsObject()									'#����/G���� ������ | 1~99999���� �Է°���, �ɼǵ�Ͻ� �ɼ�������(true)�� ������ ��� ���Ǹż����� �Է��ص� ���õǰ� �ɼ��� �ջ����� ������
					obj("itemAddtionalInfo")("stock")(Chkiif(vGubun="A", "Iac", "Gmkt")) = getLimitEbayEa()
				Set obj("itemAddtionalInfo")("sellingPeriod") = jsObject()									'#����/G���� ������ | 1~99999���� �Է°���, �ɼǵ�Ͻ� �ɼ�������(true)�� ������ ��� ���Ǹż����� �Է��ص� ���õǰ� �ɼ��� �ջ����� ������
					obj("itemAddtionalInfo")("sellingPeriod")(Chkiif(vGubun="A", "Iac", "Gmkt")) = 90
					obj("itemAddtionalInfo")("managedCode") = ""& FItemid &""						'#�Ǹ��� ��ǰ�ڵ�
				Set obj("itemAddtionalInfo")("recommendedOpts") = jsObject()
					obj("itemAddtionalInfo")("recommendedOpts")("type") = 0							'#��õ�ɼ� ��뿩�� | 0 �ɼǹ̻��, 1 ������(�ִ�20��), 2 2��������
					obj("itemAddtionalInfo")("recommendedOpts")("isStockManage") = false			'�ɼ������� | ��õ�ɼ� ����� �� �ʼ�
					obj("itemAddtionalInfo")("recommendedOpts")("independent") = null				'(������/������ ����)
					obj("itemAddtionalInfo")("recommendedOpts")("combination") = null				'(������/������ ����)
					obj("itemAddtionalInfo")("inventoryCode") = null								'(G���Ͽ�)G���� �κ��丮 �ڵ�
				Set obj("itemAddtionalInfo")("sellerShop") = jsObject()
					obj("itemAddtionalInfo")("sellerShop")("catCode") = FtenCateLarge & FtenCateMid & FtenCateSmall	'#�Ǹ��� ī�װ��ڵ�
					obj("itemAddtionalInfo")("sellerShop")("catName") = FtenCateSmall				'#�Ǹ��� ī�װ���
					obj("itemAddtionalInfo")("sellerShop")("brandCode") = FMakerId					'#�Ǹ��� �귣���ڵ�
					obj("itemAddtionalInfo")("sellerShop")("brandName") = FMakerName				'#�Ǹ��� �귣���
					obj("itemAddtionalInfo")("expiryDate") = null									'��ȿ��
					obj("itemAddtionalInfo")("manufacturedDate") = null								'������
				Set obj("itemAddtionalInfo")("origin") = jsObject()
					obj("itemAddtionalInfo")("origin")("goodsType") = 1								'#��������ǰ Ÿ�� | 0 ������ǥ�ô��ƴ�(��ǰ�̿�), 1 �󼼼�������, 2 ����ǰ, 3 ��깰, 4 ���깰
					obj("itemAddtionalInfo")("origin")("type") = 5									'#���������� Ÿ�� | 0 ����, 1 ������, 2 ���Ի�, 5 ��Ÿ | �󼼼��������ϰ�� 0~5 �� ����, �� ��ǰ��ȸ�� �󼼼��������� 0���� ������
					obj("itemAddtionalInfo")("origin")("code") = null								'���������� �ڵ�
					obj("itemAddtionalInfo")("origin")("isMultipleOrigin") = false					'#���������� ���� | true ���������� ��ǰ, false ���Ͽ����� ��ǰ
					obj("itemAddtionalInfo")("capacity") = null
'				Set obj("itemAddtionalInfo")("capacity") = jsObject()
'					obj("itemAddtionalInfo")("capacity")("vol") = null								'(���ǻ�ǰ��)�뷮/�԰� ��
'					obj("itemAddtionalInfo")("capacity")("unit") = null								'(���ǻ�ǰ��)�뷮/�԰� ����

				Set obj("itemAddtionalInfo")("shipping") = jsObject()
					obj("itemAddtionalInfo")("shipping")("type") = 1								'#��۹�� Ÿ�� �Է� | G������ ������ 1���� ��밡�� / ���� 3�� ���ý� �Ϲݿ���, ������ �湮������ ���� �ʿ� | 1 �ù����, 2 ȭ�����, 3 �Ǹ����������
					obj("itemAddtionalInfo")("shipping")("companyNo") = 10013						'#�ù���ڵ� | 10013 CJ�������
					Set obj("itemAddtionalInfo")("shipping")("policy") = jsObject()
						obj("itemAddtionalInfo")("shipping")("policy")("placeNo") = 210824			'#��������ȣ
						obj("itemAddtionalInfo")("shipping")("policy")("feeType") = 1				'#��ۺ� Ÿ��
						Set obj("itemAddtionalInfo")("shipping")("policy")("bundle") = jsObject()
							obj("itemAddtionalInfo")("shipping")("policy")("bundle")("deliveryTmplId") = 2356837 '#������ۺ���å��ȣ
						Set obj("itemAddtionalInfo")("shipping")("policy")("each") = jsObject()
							obj("itemAddtionalInfo")("shipping")("policy")("each")("feeType") = 0		'��ǰ����ۺ� Ÿ�� | ���� �������� �ʾ� ������ 0������ �Է� | 0 ������ۺ���, 1 ����, 2 ����, 3 ���Ǻι���, 4 ����������
							obj("itemAddtionalInfo")("shipping")("policy")("each")("feePayType") = 0	'��ǰ����ۺ����ҹ�� | �������� �������������� �Է�(��������)
							obj("itemAddtionalInfo")("shipping")("policy")("each")("fee") = 0			'��ǰ����ۺ�ݾ� | ������ ��� 0 �Է�(��������)
							obj("itemAddtionalInfo")("shipping")("policy")("each")("baseFee") = 0		'��ǰ����ۺ����Ǻ� �ݾ� (��������)
					Set obj("itemAddtionalInfo")("shipping")("returnAndExchange") = jsObject()
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("addrNo") = 490970			'(��ǰ�ּ�) �Ǹ����ּҹ�ȣ
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("shippingCompany") = "0008"	'��ǰ��ȯ�ù���ڵ�
						obj("itemAddtionalInfo")("shipping")("returnAndExchange")("fee") = 2500					'��ǰ/��ȯ ����ۺ�
					Set obj("itemAddtionalInfo")("shipping")("dispatchPolicyNo") = jsObject()
					'''''''''''''''''''''''�Ʒ� �� ������ ���� �߰��ؾ� ��''''''''''''''''''''''''''''''''''''''
					obj("itemAddtionalInfo")("shipping")("dispatchPolicyNo")(Chkiif(vGubun="A", "Iac", "Gmkt")) = Chkiif(vGubun="A", 587470, 587465)	'#(����/G����) �߼�Ÿ����å��ȣ
					obj("itemAddtionalInfo")("shipping")("generalPost") = null						'#(���ǿ�)�Ϲݿ��� ���� ���� �� ��ݰ���
					obj("itemAddtionalInfo")("shipping")("visitAndTake") = null						'#�湮���� ��������
					obj("itemAddtionalInfo")("shipping")("quickService") = null						'#������ ��������
				Call getEbayInfoCdParameter(obj)	'#��ǰ������� ����
				obj("itemAddtionalInfo")("isAdultProduct") = Chkiif(IsAdultItem()="Y", true, false)			'#���λ�ǰ���� | true : ���λ�ǰ, false : �Ϲݻ�ǰ
				obj("itemAddtionalInfo")("isYouthNotAvailable") = Chkiif(IsAdultItem()="Y", true, false)	'#û�ҳⱸ�źҰ����� | ��ǰ�̹����� ���⿩�� | true : û�ҳⱸ�źҰ���ǰ, false : �Ϲݻ�ǰ
				obj("itemAddtionalInfo")("isVatFree") = Chkiif(FVatInclude="N", true, false)				'#�ΰ��� ���� | true : �鼼��ǰ, false : ������ǰ
				Call getEbayCertInfoParameter(obj)	'#��ǰ������� ����
				Call getEbayImageParameter(obj)		'#��ǰ�̹��� ����
				obj("itemAddtionalInfo")("weight") = 0												'(G���Ͽ�) ��ǰ����(����:kg)
				Call getEbayContParamToReg(obj, vGubun)
				obj("itemAddtionalInfo")("addonService") = null										'�߰����� ����
			Set obj("addtionalInfo") = jsObject()
				Set obj("addtionalInfo")("sellerDiscount") = jsObject()
					obj("addtionalInfo")("sellerDiscount")("isUse") = false							'#�Ǹ������� ��뿩�� | true ��������, false ���ι�����
					Set obj("addtionalInfo")("sellerDiscount")(Chkiif(vGubun="A", "iac", "gmkt")) = jsObject()
						obj("addtionalInfo")("sellerDiscount")(Chkiif(vGubun="A", "iac", "gmkt"))("type") = 0 '����Ÿ�� | �Ǹ������� ��뿩�� true�ϰ�� �ʼ� 0 ������, 1 ����, 2 ����
				Set obj("addtionalInfo")("siteDiscount") = jsObject()
					obj("addtionalInfo")("siteDiscount")(Chkiif(vGubun="A", "iac", "gmkt")) = true		'#G����/���ǿ��� �δ��ϴ� ����Ʈ ������ �������� ���� | true ����, false ������
					obj("addtionalInfo")("gift") = null
					Set obj("addtionalInfo")("pcs") = jsObject()
						obj("addtionalInfo")("pcs")("isUse") = true									'#���ݺ񱳻���Ʈ ���⿩�� | true ���(�����), false �����������(�̳���)
						If vGubun="A" Then
						obj("addtionalInfo")("pcs")("isUseIacPcsCoupon") = false					'#(���ǿ�)���ݺ񱳻���Ʈ �������뿩��
						Else
						obj("addtionalInfo")("pcs")("isUseGmkPcsCoupon") = false					'#(G���Ͽ�)���ݺ񱳻���Ʈ �������뿩�� | G������ �ѹ� �����ϸ� ����Ұ�(������������)
						End If
					Set obj("addtionalInfo")("overseaSales") = jsObject()
						obj("addtionalInfo")("overseaSales")("isAgree") = false						'#(G���Ͽ�)�ؿ��Ǹſ��� | true ����, false �������

'		response.write obj.jsString
'		response.end
		getEbayItemRegParameter = obj.jsString
	End Function
End Class

Class CEbay
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
	Public Sub getAuctionNotRegOneItem
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
            addSql = addSql & " WHERE (optCnt-optNotSellCnt < 1) "
			addSql = addSql & " OR (optAddCNT > 0) "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.auctionStatCD,-9) as auctionStatCD "
		strSql = strSql & "	, UC.socname_kor, am.SDCategoryCode, am.cateCode "
		strSql = strSql & "	, isNull(c.isbn13, '') as isbn13 "
		strSql = strSql & "	, CONVERT(VARCHAR(10), isNull(sellSTDate, getdate()), 23) as sellSTDate "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_ebay_cate_mapping "
		strSql = strSql & "		WHERE gubun = 'A' "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ebay_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small and gubun='A' "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_auction1010_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'�ö��/ȭ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_auction1010_regItem WHERE auctionStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CEbayItem
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
			If (IsNULL(rsget("basicImage600")) or (rsget("basicImage600")="")) Then
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
			ELSE
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
			End If
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FAuctionStatCD		= rsget("auctionStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")

				FOneItem.FSDCategoryCode	= rsget("SDCategoryCode")
				FOneItem.FcateCode			= rsget("cateCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FIsbn13 			= rsget("isbn13")
'				FOneItem.FSellSTDate		= rsget("sellSTDate")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub get11stEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
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
		strSql = strSql & "	, m.st11GoodNo, m.st11price, m.st11SellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.cateCode, isNULL(m.st11StatCD,-9) as st11StatCD, tm.safeDiv, tm.isNeed, tm.depth1Code "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv = '21' "
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.st11GoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
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
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSt11GoodNo		= rsget("st11GoodNo")
				FOneItem.FSt11price			= rsget("st11price")
				FOneItem.FSt11SellYn		= rsget("st11SellYn")

				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If

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
	            FOneItem.Fregitemname   	= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
                FOneItem.FDepthCode			= rsget("depthCode")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

End Class

'11���� ��ǰ�ڵ� ���
Function get11stGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 st11goodno FROM db_etcmall.dbo.tbl_11st_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		get11stGoodno = rsget("st11goodno")
	End If
	rsget.Close
End Function

'11���� ��ǰ�ڵ�/��ǰ�� ���
Function get11stGoodno2(iitemid, ist11goodno, byref MustPrice)

	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, saleyn, tmpPrice, vdeliverytype, ispecialPrice
	Dim GetTenTenMargin, st11goodno

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.buycash, i.orgprice, i.sailyn, r.st11goodno, i.deliverytype, isnull(mi.mustPrice, 0) as specialPrice "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as mi "
	strSql = strSql & " 	on i.itemid = mi.itemid "
	strSql = strSql & " 	and mi.mallgubun = '11st1010' "
	strSql = strSql & " 	and (GETDATE() >= mi.startDate and GETDATE() <= mi.endDate ) "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		saleyn		= rsget("sailyn")
		st11goodno	= rsget("st11goodno")
		vdeliverytype = rsget("deliverytype")
		ispecialPrice = rsget("specialPrice")
	Else
		get11stGoodno2 = ""
		Exit Function
		response.end
	End If
	rsget.close

	If ispecialPrice <> "0" Then
		tmpPrice = ispecialPrice
	Else
		GetTenTenMargin = CLng((10000 - buycash / sellcash * 100 * 100) / 100)
	'	If (vdeliverytype = 2) OR (vdeliverytype = 9) Then
	'		If (GetTenTenMargin < CMAXMARGIN) OR (saleyn = "Y" AND sellcash < 10000) Then
	'			tmpPrice = orgprice
	'		Else
	'			tmpPrice = sellcash
	'		End If
	'	Else
			If (GetTenTenMargin < CMAXMARGIN) Then
				tmpPrice = orgprice
			Else
				tmpPrice = sellcash
			End If
	'	End If
	End If
	MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	ist11goodno = st11goodno
End Function

'11���� ��ǰ�ڵ�, �ɼ� �� ���
Function get11stGoodno3(iitemid, ist11goodno, byref opCnt)
	Dim strSql, st11goodno, optioncnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.optioncnt, r.st11goodno "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		opCnt		= rsget("optioncnt")
		ist11goodno	= rsget("st11goodno")
	Else
		get11stGoodno3 = ""
		Exit Function
		response.end
	End If
	rsget.close
End Function

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
