<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "sabangnet"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST sabangnetAPIURL = "http://r.sabangnet.co.kr"
CONST sabangnetID = "tenbyten"
CONST sabangnetAPIKEY = "PTxNV3d9CXPXBNu60X72EbSNYTJd5955b"
CONST CDEFALUT_STOCK = 999
CONST wapiURL = "http://wapi.10x10.co.kr"

Class CSabangnetItem
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
	Public FIcon1Image
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public FSafetydiv
	Public Fitemcontent
	Public FSabangnetStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FMaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public FItemsize
	Public FItemsource
	Public FBrandCode
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public FSabangnetGoodNo
	Public FSabangnetprice
	Public FSabangnetSellYn
	Public FMayLimitSoldout
	Public FMwdiv

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(�ٹ�), 4,5(����)
			IsFreeBeasong = True
		End If

		If (FdeliveryType=9) Then														'��ü����
			IsFreeBeasong = False
		End If
		If (MustPrice >= 50000) Then IsFreeBeasong = True
    End Function

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
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
		Dim GetTenTenMargin, sqlStr, specialPrice, tmpPrice, vBigPrice, vSmallPrice, ownItemCnt
		specialPrice = 0
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

		If specialPrice <> 0 Then
			MustPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			MustPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
			MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
		End If
	End Function

	'// ���� �Ǹſ��� ��ȯ
	Public Function getSabangnetSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getSabangnetSellYn = "Y"
			Else
				getSabangnetSellYn = "N"
			End If
		Else
			getSabangnetSellYn = "N"
		End If
	End Function

	Public Function getSourcearea()
		Dim arrAreaName, i
		arrAreaName = Array("America", "Australia", "Belgium", "Brazil", "Chile", "CHINA", "ITALY", "KOREA", "Mexico", "Norway", "Thailand", "���׸���", "������", "�׷�����(������)", "�׸���", "��Ÿ����", "����������", "��������ī��ȭ��", "�״�����", "����", "�븣����", "��������", "��ī���", "�븸", "���ѹα�", "����ũ", "���̴�ī", "���̴�ī��ȭ��", "����", "�����", "��Ʈ���", "���þ�", "���þư�ȭ��", "���ٳ�", "�縶�Ͼ�", "�����", "�����ƴϾ�", "���ٰ���ī��", "��ī��", "�����̽þ�", "����������", "�߽���", "�����", "�𸮼Ž�", "���ٺ��", "������", "��Ÿ", "����", "�̱�", "�̱�/�Ϻ�", "�̱�OEM", "�̾Ḷ", "�ٷ���", "��۶󵥽�", "���׼�����", "��Ʈ��", "���⿡", "�����Ͼ�", "���տ�����", "�������", "����", "�Ұ�����", "�����", "����ƶ���", "�󼼼�������", "���װ�", "�������", "���Ի�", "������ī", "������", "������", "����Ʋ����", "������", "���ι�Ű��", "���κ��Ͼ�", "�̰�����", "�ƶ����̷���Ʈ", "�ƶ����̸�Ʈ", "�Ƹ��޴Ͼ�", "�Ƹ���Ƽ��", "����Ƽ", "���Ϸ���", "������ī", "�˹ٴϾ�", "������Ͼ�", "���⵵��", "����ٵ���", "����", "����Ʈ���ϸ���", "����Ʈ����", "�µζ�", "�ܱ���", "�丣��", "�찣��", "������", "���Ű��ź", "��ũ���̳�", "�����", "��������(EU)", "�̵���Ǿ�", "�̶�ũ", "�̶�", "�̽���", "����Ʈ", "��Ż����", "���¸�", "�ε�", "�ε��׽þ�", "�ε��׽þ�OEM", "�ε��", "�Ϻ�", "�Ϻ�/�±�", "�߱�", "�߱�/�븸", "�߱�/�����̽þ�", "�߱�/�̾Ḷ", "�߱�/��Ʈ��", "�߱�/�ε�", "�߱�/�ε��׽þ�", "�߱�/�±�", "�߱�/�ʸ���", "�߱�OEM", "�߱�����", "�߱��ܺ�������", "����Ƽ", "ü��", "ĥ��", "į�����", "ĳ����", "�ɳ�", "�̸����Ͼ�", "�ݷҺ��", "�����Ʈ", "ũ�ξ�Ƽ��", "Ÿ�̿�", "Ÿ�Ϸ���", "�±�", "��Ű", "Ƣ����", "�Ķ����", "��Ű��ź", "���", "��������", "������", "������", "������/�̱�", "������/�߱�", "�ɶ���", "�ʸ���", "�ѱ�/�߱�", "�ѱ�/�߱�/�̱�", "�밡��", "ȣ��", "ȫ��")

		For i =0 To Ubound(arrAreaName)
			If Trim(arrAreaName(i)) = Trim(FSourcearea) Then
				getSourcearea = Trim(arrAreaName(i))
				Exit For
			End If
		Next

		If FSourcearea = "�ѱ�" Then
			getSourcearea = "���ѹα�"
		End If

		If getSourcearea = "" Then
			getSourcearea = "��Ÿ����"
		End If
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
		If trim(Fkeywords) = "" Then Exit Function
		Fkeywords  = replace(Fkeywords,"%", "")
		Fkeywords  = replace(Fkeywords,"/", ",")
		Fkeywords  = replace(Fkeywords,".", "")
		Fkeywords  = replace(Fkeywords,"+", "")
		Fkeywords  = replace(Fkeywords,"_", "")
		Fkeywords  = replace(Fkeywords,"(", "")
		Fkeywords  = replace(Fkeywords,")", "")
		Fkeywords  = replace(Fkeywords,"&", "")
		Fkeywords  = replace(Fkeywords,";", "")
		Fkeywords  = replace(Fkeywords,"#", "")
		Fkeywords  = replace(Fkeywords,"'", "")
		Fkeywords  = replace(Fkeywords,"[", "")
		Fkeywords  = replace(Fkeywords,"]", "")
		Fkeywords  = replace(Fkeywords,":", "")
		Fkeywords  = replace(Fkeywords,"\", "")

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
'rw getItemKeyword
'response.end
	End Function

    Public Function getItemNameFormat()
		Dim buf
		If application("Svr_Info") = "Dev" Then
			FItemName = "TEST��ǰ "&FItemName
		End If

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
		getItemNameFormat = buf
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

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
		i = 0
		If Flimityn = "Y" Then
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
				End If
				rsget.Close
			End If
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getSabangnetContParamToReg()
		Dim strRst, strSQL, infoContRst
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style><br>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_iPark.jpg'></p><br>"
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

		strSQL = ""
		strSQL = strSQL & " SELECT c.infoCd, c.infoItemName, "
		strSQL = strSQL & " CASE WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035' "
		strSQL = strSQL & " 	WHEN c.infotype='T' and c.infoItemName = 'ǰ����������' THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' "
		strSql = strSql & " 	WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '��ǰ �� ����' " & vbcrlf
		strSQL = strSQL & " ELSE F.infocontent END AS infocontent "
		strSQL = strSQL & " FROM db_item.dbo.tbl_item_contents IC "
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_infoCode c ON ic.infoDiv=c.infoDiv "
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON F.infoCd = c.infoCd and F.itemid='"& Fitemid &"' "
		strSQL = strSQL & " WHERE IC.itemid='"& Fitemid &"' "
		strSQL = strSQL & " ORDER BY convert(int, c.infoCd) ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			infoContRst = ""
			infoContRst = infoContRst & "<table align=""center"">"
			infoContRst = infoContRst & "	<colgroup>"
			infoContRst = infoContRst & "		<col style=""width: 30%;"">"
			infoContRst = infoContRst & "		<col style=""width: 70%;"">"
			infoContRst = infoContRst & "	</colgroup>"
			infoContRst = infoContRst & "	<tbody>"
			Do until rsget.EOF
				infoContRst = infoContRst & "<tr>"
				infoContRst = infoContRst & "	<td scope=""row"">"&rsget("infoItemName")&"</td>"
				infoContRst = infoContRst & "	<td>"&rsget("infocontent")&"</td>"
				infoContRst = infoContRst & "</tr>"
				rsget.MoveNext
			Loop
			infoContRst = infoContRst & "	</tbody>"
			infoContRst = infoContRst & "</table>"
			strRst = strRst & infoContRst
		End If
		rsget.Close

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_sabangnet.jpg"">")
		getSabangnetContParamToReg = strRst
	End Function

	Public Function getSabangnetOptParamtoREG()
		Dim buf, sqlStr, i, tmpVAL, limitYCnt, limitNCnt
		Dim vitemoption, voptionname, voptsellyn, voptaddprice, voptLimit, optStatus
    	buf = ""
    	tmpVAL = ""
    	limitYCnt = 0
    	limitNCnt = 0

		If FOptionCnt > 0 Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemoption, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
			sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					vitemoption 		= rsget("itemoption")
					voptionname 		= Replace(rsget("optionname"), ",", "_")
					voptsellyn 			= rsget("optsellyn")
					voptaddprice		= rsget("optaddprice")
					voptLimit			= rsget("optLimit")
					voptLimit = voptLimit-5
					If (voptLimit < 1) Then voptLimit = 0
					If (FLimitYN <> "Y") Then voptLimit = CDEFALUT_STOCK

					If ((voptsellyn <> "Y") OR (voptLimit = 0)) Then
						optStatus = "004"			'004 : ǰ��
					Else
						optStatus = "002"			'002 : ������
					End If

					tmpVAL = tmpVAL & voptionname &"^^"& voptLimit &"^^"& voptaddprice &"^^"& vitemoption &"^^EA^^"& optStatus & ","
					If (voptLimit = 0) Then
						limitNCnt = limitNCnt + 1
					Else
						limitYCnt = limitYCnt + 1
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
			tmpVAL = RightCommaDel(tmpVAL)

			If FOptioncnt > 0 Then
				If limitYCnt = 0 Then
					FMayLimitSoldout = "Y"
				Else
					FMayLimitSoldout = "N"
				End If
			End If

			buf = buf & "		<CHAR_1_NM><![CDATA[�ɼ�]]></CHAR_1_NM>"
			buf = buf & "		<CHAR_1_VAL><![CDATA["&tmpVAL&"]]></CHAR_1_VAL>"
		Else
			buf = buf & "		<CHAR_1_NM><![CDATA[��ǰ]]></CHAR_1_NM>"
			If Flimityn = "Y" Then
				voptLimit = FLimitNo - FLimitSold -5
			Else
				voptLimit = CDEFALUT_STOCK
			End If

			If voptLimit < 1 Then
				voptLimit = 0
			End If

			buf = buf & "		<CHAR_1_VAL><![CDATA[��ǰ^^"&voptLimit&"]]></CHAR_1_VAL>"
		End If
		buf = buf & "		<CHAR_2_NM><![CDATA[]]></CHAR_2_NM>"
		buf = buf & "		<CHAR_2_VAL><![CDATA[]]></CHAR_2_VAL>"
		getSabangnetOptParamtoREG = buf
	End Function

	Public Function getSabangnetAddImageParam()
		Dim strRst, strSQL, i
		If application("Svr_Info")="Dev" Then
			'FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
			FbasicImage = "http://webimage.10x10.co.kr/image/basic/71/B000712763-10.jpg"
		End If
		strRst = ""
		strRst = strRst & "		<IMG_PATH><![CDATA["&FbasicImage&"]]></IMG_PATH>"					'#��ǥ�̹��� | �� : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg
		strRst = strRst & "		<IMG_PATH1><![CDATA["&FbasicImage&"]]></IMG_PATH1>"					'#���ո�(JPG)�̹��� | �� : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg  (���ո�(JPG)�̹��� (500x500 ~ 700x700))
		strRst = strRst & "		<IMG_PATH2><![CDATA[]]></IMG_PATH2>"								'�ΰ��̹���2
		strRst = strRst & "		<IMG_PATH3><![CDATA["&FIcon1Image&"]]></IMG_PATH3>"					'#�ΰ��̹���3 | �� : http://gs4333.CO.KR/product_image/a0000769/200907/image20_700.jpg  (11��������̹��� (300*300))
		strRst = strRst & "		<IMG_PATH4><![CDATA[]]></IMG_PATH4>"								'�ΰ��̹���4
		strRst = strRst & "		<IMG_PATH5><![CDATA[]]></IMG_PATH5>"								'�ΰ��̹���5
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					If (isNull(rsget("addimage_600")) OR rsget("addimage_600") = "") Then
						strRst = strRst & "		<IMG_PATH"&i+5&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"]]></IMG_PATH"&i+5&">"	'�ΰ��̹��� 6~10 | ���θ� �߰��̹���(1~5)
					Else
						strRst = strRst & "		<IMG_PATH"&i+5&"><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")&"]]></IMG_PATH"&i+5&">"	'�ΰ��̹��� 6~10 | ���θ� �߰��̹���(1~5)
					End If
				End If
				rsget.MoveNext
				If i>=5 Then Exit For
			Next
		End If
		rsget.Close
		getSabangnetAddImageParam = strRst
	End Function

	Public Function getSabangnetCertInfoToReg
		Dim buf, strSql, safetyDiv, safetyId, certNum, certOrganName, certmakerName, isRegCert, certDiv
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN '�����ǰ_��������' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN '�����ǰ_����Ȯ�νŰ�' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN '�����ǰ_���������ռ�Ȯ��' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN '��Ȱ��ǰ_��������' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN '��Ȱ��ǰ_����Ȯ�νŰ�' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN '��Ȱ��ǰ_���������ռ�Ȯ��' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN '�����ǰ_��������' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN '�����ǰ_����Ȯ�νŰ�' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN '�����ǰ_���������ռ�Ȯ��' end as safetyId "
		strSql = strSql & " , t.certNum, f.certOrganName, f.makerName, f.certDiv "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			certNum			= rsget("certNum")
			certOrganName	= rsget("certOrganName")
			certmakerName	= rsget("makerName")
			certDiv			= rsget("certDiv")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		buf = ""
		buf = buf & "		<CERTNO><![CDATA["&certNum&"]]></CERTNO>"								'������ȣ | �����ǰ, ���� ������ǰ �� �����˻縦 ���ľ� �ϴ� ��ǰ�� ��� �ش������� �ο��� ������ȣ�� �Է��մϴ�
		buf = buf & "		<AVLST_DM></AVLST_DM>"													'������ȿ ������ | ����8�ڸ� �Է��ϼ��� ��:20100401
		buf = buf & "		<AVLED_DM></AVLED_DM>"													'������ȿ �������� | ����8�ڸ� �Է��ϼ��� ��:20100401
		buf = buf & "		<ISSUEDATE></ISSUEDATE>"												'�߱����� | ����8�ڸ� �Է��ϼ��� ��:20100401
		buf = buf & "		<CERTDATE></CERTDATE>"													'�������� | ����8�ڸ� �Է��ϼ��� ��:20100401
		buf = buf & "		<CERT_AGENCY><![CDATA["&certOrganName&"]]></CERT_AGENCY>"				'������� | �� : �ѱ��������������
		buf = buf & "		<CERTFIELD><![CDATA["&certDiv&"]]></CERTFIELD>"							'�����о� | �� : �԰�����
		getSabangnetCertInfoToReg = buf
	End Function

	' Public Function getSabangnetCertInfoToReg
	' 	Dim buf, safetydivName, certNo, ssgCERTFIELD
	' 	If (FSafetyyn = "Y") and (Trim(FSafetyNum) <> "") Then
	' 		certNo = Trim(FSafetyNum)
	' 		Select Case FSafetydiv
	' 			Case "10"
	' 				safetydivName = "������������(KC��ũ)"
	' 				ssgCERTFIELD = "��������_��������"
	' 			Case "20"
	' 				safetydivName = "�����ǰ ��������"
	' 				ssgCERTFIELD = "��������"
	' 			Case "30"
	' 				safetydivName = "KPS �������� ǥ��"
	' 				ssgCERTFIELD = "��������_��������"
	' 			Case "40"
	' 				safetydivName = "KPS �������� Ȯ�� ǥ��"
	' 				ssgCERTFIELD = "��������_����Ȯ��"
	' 			Case "50"
	' 				safetydivName = "KPS ��� ��ȣ���� ǥ��"
	' 				ssgCERTFIELD = "��������_��������"
	' 		End Select
	' 		'##SSG�� �Է��ؾ� �Ǵ� ���̹�
	' 		'��������_��������
	' 		'��������_����Ȯ��
	' 		'��������_���������ռ�
	' 		'��������
	' 		'���ؿ��
	' 	End If

	' 	If ssgCERTFIELD = "" Then
	' 		ssgCERTFIELD = "�����ǰ^^��������"
	' 		certNo = "����^^����"
	' 	End If

	' 	buf = ""
	' 	buf = buf & "		<CERTNO><![CDATA["&certNo&"]]></CERTNO>"								'������ȣ | �����ǰ, ���� ������ǰ �� �����˻縦 ���ľ� �ϴ� ��ǰ�� ��� �ش������� �ο��� ������ȣ�� �Է��մϴ�
	' 	buf = buf & "		<AVLST_DM></AVLST_DM>"													'������ȿ ������ | ����8�ڸ� �Է��ϼ��� ��:20100401
	' 	buf = buf & "		<AVLED_DM></AVLED_DM>"													'������ȿ �������� | ����8�ڸ� �Է��ϼ��� ��:20100401
	' 	buf = buf & "		<ISSUEDATE></ISSUEDATE>"												'�߱����� | ����8�ڸ� �Է��ϼ��� ��:20100401
	' 	buf = buf & "		<CERTDATE></CERTDATE>"													'�������� | ����8�ڸ� �Է��ϼ��� ��:20100401
	' 	buf = buf & "		<CERT_AGENCY><![CDATA[]]></CERT_AGENCY>"								'������� | �� : �ѱ��������������
	' 	buf = buf & "		<CERTFIELD><![CDATA["&ssgCERTFIELD&"]]></CERTFIELD>"					'�����о� | �� : �԰�����
	' 	getSabangnetCertInfoToReg = buf
	' End Function

	Public Function getSabangnetItemInfoCdToReg
		Dim strSql, buf, lp
		Dim mallinfoCd, infoContent, rsMallinfoDiv
		strSql = ""
		strSql = strSql & " SELECT TOP 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCd='00001') AND (IC.safetyyn= 'Y') THEN IC.safetyNum "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035'  "
		strSql = strSql & " 	 WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '��ǰ �� ����' " & vbcrlf
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv  "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd  "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"'  "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = 'sabangnet' and IC.itemid='"&FItemID&"'"
		strSql = strSql & " ORDER BY convert(int, mallinfocd) ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			rsMallinfoDiv = rsget("mallinfoDiv")
			If rsMallinfoDiv = "47" Then
				rsMallinfoDiv = "36"
			ElseIf rsMallinfoDiv = "" Then
				rsMallinfoDiv = "37"
			End If

			buf = ""
			buf = buf & "		<PROP_EDIT_YN>Y</PROP_EDIT_YN>"										'�Ӽ��������� | "�Ӽ����� �������θ� Y or N�� �Է��մϴ�. Y�Է½� �Ӽ�����(�Ӽ��з��ڵ�, �Ӽ���)�� ���� ó���մϴ�."
			buf = buf & "		<PROP1_CD>0"& rsMallinfoDiv &"</PROP1_CD>"					'�Ӽ��з��ڵ� | "�Ӽ��з��ڵ带 ���� 3�ڸ� �������� �Է��մϴ�. �Ӽ��з��ڵ�� ��ǰ�Ӽ��ڵ� ��ȸ API�� ���� ��ǰ����ȭ���� �Ӽ��з�ǥ�� �����Ͻñ� �ٶ��ϴ�. ��: �Ƿ��� 001�� �Է��մϴ�."
			Do until rsget.EOF
				infoContent = rsget("infocontent")
				mallinfocd = rsget("mallinfocd")
				buf = buf & "		<PROP_VAL"&mallinfocd&"><![CDATA["&infoContent&"]]></PROP_VAL"&mallinfocd&">"	'�Ӽ��� | "�Ӽ��з��ڵ忡 ���� �Ӽ�����1�� �Ӽ��� �ش��ϴ� �Ӽ����� �Է��մϴ�. �Ӽ���(1 ~ 20)��  �Է¼������ ó���ǹǷ�, �Ӽ������� �����Ͻñ� �ٶ��ϴ�.(�Ӽ����� ���� ���, �������� ó���Ͻñ� �ٶ��ϴ�.) �� : �Ƿ� 001�� �Ӽ���1�� ��ǰ �����̸�, �Ӽ���1�� ��,���Ϸ� ���� �Է��մϴ�."
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If (session("ssBctID")="kjy8517") and FItemId = "1882712" Then
			buf = ""
			buf = buf & "		<PROP_EDIT_YN>Y</PROP_EDIT_YN>"				'�Ӽ��������� | "�Ӽ����� �������θ� Y or N�� �Է��մϴ�. Y�Է½� �Ӽ�����(�Ӽ��з��ڵ�, �Ӽ���)�� ���� ó���մϴ�."
			buf = buf & "		<PROP1_CD>008</PROP1_CD>"					'�Ӽ��з��ڵ� | "�Ӽ��з��ڵ带 ���� 3�ڸ� �������� �Է��մϴ�. �Ӽ��з��ڵ�� ��ǰ�Ӽ��ڵ� ��ȸ API�� ���� ��ǰ����ȭ���� �Ӽ��з�ǥ�� �����Ͻñ� �ٶ��ϴ�. ��: �Ƿ��� 001�� �Է��մϴ�."
			buf = buf & "		<PROP_VAL1><![CDATA[�������� ����]]></PROP_VAL1>"
			buf = buf & "		<PROP_VAL2><![CDATA[�������� ����]]></PROP_VAL2>"
			buf = buf & "		<PROP_VAL3><![CDATA[�������� ����]]></PROP_VAL3>"
			buf = buf & "		<PROP_VAL4><![CDATA[�������� ����]]></PROP_VAL4>"
			buf = buf & "		<PROP_VAL5><![CDATA[�������� ����]]></PROP_VAL5>"
			buf = buf & "		<PROP_VAL6><![CDATA[�������� ����]]></PROP_VAL6>"
			buf = buf & "		<PROP_VAL7><![CDATA[�������� ����]]></PROP_VAL7>"
			buf = buf & "		<PROP_VAL8><![CDATA[�������� ����]]></PROP_VAL8>"
			buf = buf & "		<PROP_VAL9><![CDATA[�������� ����]]></PROP_VAL9>"
			buf = buf & "		<PROP_VAL10><![CDATA[�������� ����]]></PROP_VAL10>"
			buf = buf & "		<PROP_VAL11><![CDATA[�������� ����]]></PROP_VAL11>"
			buf = buf & "		<PROP_VAL12><![CDATA[�������� ����]]></PROP_VAL12>"
			buf = buf & "		<PROP_VAL13><![CDATA[�������� ����]]></PROP_VAL13>"
			buf = buf & "		<PROP_VAL14><![CDATA[�������� ����]]></PROP_VAL14>"
			buf = buf & "		<PROP_VAL15><![CDATA[�������� ����]]></PROP_VAL15>"
		End If
		getSabangnetItemInfoCdToReg = buf
	End Function

	'��ǰ ��� XML
	Public Function getSabangnetItemRegParameter(isReg, ichgSellyn)
		Dim strRst, tmpStatus, vMwdiv
		If isReg = False Then
			'4(����ǰ��)�� ������ �� ���� ����� ���θ��� �ִٸ� ���� ���� ������ �� �� �ִ���
			'3(�Ͻ�����)���� ������ �ڵ�ȭ ���� ���� �Ǹ�/ǰ���� �Ը��� �°� ���氡�� �亯 ����
			tmpStatus = 2
		Else
			If ichgSellyn = "Y" Then
				tmpStatus = 2
			Else
				tmpStatus = 3
			End If
		End If

		Select Case FMwdiv
			Case "M"	vMwdiv = "3"
			Case Else	vMwdiv = "1"
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#���� �α��� ���̵�
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#���ݿ��� �߱� ���� ����Ű
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#�������� | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'��ü�ڵ� ��ȯ���� | ��� ������ ����� ��ü�ڵ� ǥ���� (Y : ��ȯ, NULL : ����)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<GOODS_NM><![CDATA["&getItemNameFormat&"]]></GOODS_NM>"				'#��ǰ�� | �ѱ۱��� 50�ڸ����� ��밡���ϸ� , HTML �±� ����� �Ұ��մϴ�.
		strRst = strRst & "		<GOODS_KEYWORD></GOODS_KEYWORD>"									'��ǰ��� | ������ ��ǰ�����ν� �ù���� ��°� ���� ������� ���� �ν��� ���Ͽ� ����� �� �ֽ��ϴ�. ( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<MODEL_NM></MODEL_NM>"												'�𵨸� | ��ǰ�� �𵨸��� ��Ȯ�� �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<MODEL_NO></MODEL_NO>"												'��No | ��ǰ�� ��No.�� ��Ȯ�� �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<BRAND_NM><![CDATA["&chkIIF(trim(FSocname_kor)="" or isNull(FSocname_kor),"��ǰ���� ����",FSocname_kor)&"]]></BRAND_NM>"	'�귣��� | �귣����� �����մϴ�.
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#��ü��ǰ�ڵ� | �ڻ翡�� ����ϴ� ��ǰ�ڵ带 �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<GOODS_SEARCH><![CDATA["&RightCommaDel(Trim(getItemKeyword()))&"]]></GOODS_SEARCH>"	'����Ʈ�˻��� | ���θ��� ��ǰ���� ���۽� ���� ����Ʈ�˻�� �޸�(,)�� �����Ͽ� �Է��մϴ�.( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<GOODS_GUBUN><![CDATA["&vMwdiv&"]]></GOODS_GUBUN>"					'#��ǰ���� | ��ǰ�� ������ ���ڷ� �Է��մϴ�. 1.��Ź��ǰ 2.������ǰ 3.���Ի�ǰ 4.������ǰ
		strRst = strRst & "		<CLASS_CD1><![CDATA["&FTenCateLarge&"]]></CLASS_CD1>"				'#��з��ڵ� | ���ݿ� ��ϵ� ��з��ڵ带 �Է��մϴ�.( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<CLASS_CD2><![CDATA["&FTenCateMid&"]]></CLASS_CD2>"					'#�ߺз��ڵ� | ���ݿ� ��ϵ� �ߺз��ڵ带 �Է��մϴ�.( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<CLASS_CD3><![CDATA["&FTenCateSmall&"]]></CLASS_CD3>"				'#�Һз��ڵ� | ���ݿ� ��ϵ� �Һз��ڵ带 �Է��մϴ�.( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<CLASS_CD3><![CDATA[]]></CLASS_CD3>"								'���з��ڵ� | ���ݿ� ��ϵ� ���з��ڵ带 �Է��մϴ�.( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<PARTNER_ID><![CDATA[]]></PARTNER_ID>"								'����óID | ����ó�� ID�� �����մϴ�.(��/�ҹ��� ��Ȯ�� �����ؾ� ��)
		strRst = strRst & "		<DPARTNER_ID><![CDATA[]]></DPARTNER_ID>"							'����óID | ����ó�� ID�� �����մϴ�.(��/�ҹ��� ��Ȯ�� �����ؾ� ��)
		strRst = strRst & "		<MAKER><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></MAKER>"	'������ | ����ȸ���� ��Ī�� ��Ȯ�� �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<ORIGIN><![CDATA["&getSourcearea&"]]></ORIGIN>"						'#������(������) | ��:�߱�,���� ������ ǥ�� �����Ͻþ� ǥ�� ����Ǿ� �ִ� ������ ������ �������ּ���. �������� ��ϵǾ� ���� �ʴ� ��� �ݼ��ͷ� ��û�Ͻñ� �ٶ��ϴ� ( "NULL" �̰ų� "���°�" �� ��� "��Ÿ" �� �Էµ�)
		strRst = strRst & "		<MAKE_YEAR><![CDATA[]]></MAKE_YEAR>"								'���꿬�� | ��ǰ�� ����� �⵵�� ���� 4�ڸ��� �Է��մϴ�. �� : 2009
		strRst = strRst & "		<MAKE_DM><![CDATA["& replace(Date(), "-", "") &"]]></MAKE_DM>"									'�������� | ��ǰ�� ������ ���ڸ� ���� 8�ڸ��� �Է��մϴ�. �� : 20100101
		strRst = strRst & "		<GOODS_SEASON>7</GOODS_SEASON>"										'#���� | ������ ������ ���ڷ� �Է��մϴ�. 1.�� 2.���� 3.���� 4.�ܿ� 5.FW 6.SS 7.�ش����  ( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<SEX>4</SEX>"														'#���౸�� | ���������� ���ڷ� �Է��մϴ�. 1.������ 2.������ 3.���� 4.�ش���� ( ��, "NULL"�̸� ��������)
		strRst = strRst & "		<STATUS>"&tmpStatus&"</STATUS>"										'#��ǰ���� | ��ǰ�� ���޻��¿� ���� �����ڵ带 �����մϴ�. 1.����� 2.������ 3.�Ͻ����� 4.����ǰ�� 5.�̻�� 6.����
		strRst = strRst & "		<DELIV_ABLE_REGION>1</DELIV_ABLE_REGION>"							'�Ǹ����� | �ǸŰ��������� ���ڷ� �Է��մϴ�. 1.���� 2.����(��������) 3.������ 4.��Ÿ
		strRst = strRst & "		<TAX_YN>"&Chkiif(Fvatinclude="Y", "1", "2")&"</TAX_YN>"				'#�������� | �������θ� ���ڷ� �Է��մϴ�. 1.���� 2.�鼼 3.�ڷ���� 4.�����
		strRst = strRst & "		<DELV_TYPE>"&Chkiif(IsFreeBeasong = True, "1", "3")&"</DELV_TYPE>"	'#��ۺ񱸺� | ��ۺ� ������ ���ڷ� �Է��մϴ�. 1.���� 2.���� 3.������ 4.����/������
		strRst = strRst & "		<DELV_COST><![CDATA["&CHKIIF(IsFreeBeasong=False,"3000","0")&"]]></DELV_COST>"	'��ۺ� | ��ۺ� ���ڷ� �Է��մϴ�. ù���ڴ� �ݵ�� '(ENTER����Key)�� �����ؾ��ϸ� ���ڻ��̿� �޸�(,)�� ���� �ȵ˴ϴ�.
		strRst = strRst & "		<BANPUM_AREA></BANPUM_AREA>"										'��ǰ������ | ����ó�� ������ ��ǰ���� �ش��ϴ� ������ �����մϴ�. �� : 1, �����ϰ�� �⺻�ּҰ� ����˴ϴ�.
		strRst = strRst & "		<GOODS_COST><![CDATA["&Clng(GetRaiseValue(FBuycash/10)*10)&"]]></GOODS_COST>"	'#���� | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & "		<GOODS_PRICE><![CDATA["& Clng(MustPrice/10)*10 &"]]></GOODS_PRICE>"	'#�ǸŰ� | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & "		<GOODS_CONSUMER_PRICE><![CDATA["&Clng(FOrgPrice/10)*10&"]]></GOODS_CONSUMER_PRICE>"	'#TAG��(�Һ��ڰ�) | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & getSabangnetOptParamtoREG()
		strRst = strRst & getSabangnetAddImageParam()
		strRst = strRst & "		<GOODS_REMARKS><![CDATA["&getSabangnetContParamToReg()&"]]></GOODS_REMARKS>"	'#��ǰ�󼼼��� | ��ǰ��(HTML)�� �����մϴ�.
		strRst = strRst & getSabangnetCertInfoToReg()
		strRst = strRst & "		<MATERIAL><![CDATA[]]></MATERIAL>"									'��ǰ���/������ | "�ؽ�ǰ�� ���� ������ ������ /(������)�� ǥ���ϸ� �߰� �Է� �� ,(�޸�)�� �����Ͽ� �߰��� ���� �������� �Է��մϴ�. �ǸŽ�ǰ �������� ��: ����/ȣ����,���/������ "
		'############################################################
		'�߿�! : STOCK_USE_YN�� Y�� ������ OPT_TYPE�� 9�� �����ؾ���
		'		STOCK_USE_YN�� N���� ������ OPT_TYPE�� 2�� ���� ����
		strRst = strRst & "		<STOCK_USE_YN><![CDATA[N]]></STOCK_USE_YN>"							'#��������뿩�� | "������ ��뿩�θ� Y or N�� �Է��մϴ�.  Y�Է½� [������] �޴����� �ش��ǰ�� ���� ��/��� �����ϸ�, ���θ��� ��ǰ������ ���������� �����˴ϴ�. N�Է½� [������] �޴����� [������(�ֹ�)] �޴��� ��밡���ϸ�, ���θ��� ��ǰ������ �������� �����˴ϴ�. ��ǰ�� ������� �Է��� [��ǰ����] >> [��ǰ�뷮����] ���� �Է°����մϴ�. "
		strRst = strRst & "		<OPT_TYPE><![CDATA[2]]></OPT_TYPE>"									'#�ɼǼ������� | "��ǰ������ ��ϵ� �ɼ��� ������ ��� ����� ���� ����ϴ� �ɼ��Դϴ�. 9: �ɼ��� ������ ������ �ʴ´�. ���� �������� �̿��ϴ� ��ü�ΰ�� �ɼ��� ������ ����� �Ǹ� ������ ������ �ִ� �ɼ��ڵ尪�� �Ҿ������ ������ �������� ū������ �߻��ϹǷ� �������� ����Ǿ� �ִ� ��ü��� �ɼ������� ������ �����ϼž� �մϴ�. �ѹ� 9�� ���õ� ��ǰ�� �ٸ� �������� �Ұ����մϴ�. 2: ��ϵ� �ɼ��� ������ ��� ����� ���� �ɼ��� �����Ѵ�. ���� 2�� ������ �Ǹ� �� ��ǰ�� ����Ǿ� �ִ� �ɼ��� ������ ��� �����ϰ� ������ �������� �ɼ��� �籸���մϴ�. �̶�, ������ �ɼ��� �������ɿ��ο� ���ؼ� 9�� �����ϼ̴ٸ�, 2�� ������ �Ұ����մϴ�. ��, ������ �ɼ��� �������ɿ��ο� 2�� �����ϼ̴ٰ� 9�� ������ �����մϴ�."
		'############################################################
		strRst = strRst & getSabangnetItemInfoCdToReg()
		strRst = strRst & "		<PACK_CODE_STR><![CDATA[]]></PACK_CODE_STR>"						'�߰���ǰ�׷��ڵ� | ���ݿ� �ԷµǾ� ������ �߰���ǰ�� �׷��� �����մϴ�. �� : G001,G004,G201 (7���� �׷��� �Է� ������)
		strRst = strRst & "		<GOODS_NM_EN><![CDATA[]]></GOODS_NM_EN>"							'���� ��ǰ�� | ���� 100�ڸ����� ��밡���ϸ� , HTML �±� ����� �Ұ��մϴ�.
		strRst = strRst & "		<GOODS_NM_PR><![CDATA[]]></GOODS_NM_PR>"							'��� ��ǰ�� | �ѱ۱��� 50�ڸ����� ��밡���ϸ� , HTML �±� ����� �Ұ��մϴ�.
		strRst = strRst & "		<GOODS_REMARKS2><![CDATA[]]></GOODS_REMARKS2>"						'�߰� ��ǰ�󼼼���_1 | ��ǰ �߰���(HTML)�� �����մϴ�. (��, "DEL" �Է½� ����� �߰��󼼼���1�� �����մϴ�.)
		strRst = strRst & "		<GOODS_REMARKS3><![CDATA[]]></GOODS_REMARKS3>"						'�߰� ��ǰ�󼼼���_2 | ��ǰ �߰���(HTML)�� �����մϴ�. (��, "DEL" �Է½� ����� �߰��󼼼���2�� �����մϴ�.)
		strRst = strRst & "		<GOODS_REMARKS4><![CDATA[]]></GOODS_REMARKS4>"						'�߰� ��ǰ�󼼼���_3 | ��ǰ �߰���(HTML)�� �����մϴ�. (��, "DEL" �Է½� ����� �߰��󼼼���3�� �����մϴ�.)
		strRst = strRst & "		<IMPORTNO><![CDATA[]]></IMPORTNO>"									'���ԽŰ��ȣ | ��ǰ ���ԽŰ��ȣ�� �����մϴ�. (12345-12-123456U)
		strRst = strRst & "		<GOODS_COST2><![CDATA[]]></GOODS_COST2>"							'����2 | ����2�� ��ǰ�۽�,�ֹ�����,��������,���� �̿���� ������, ������ ���� ���� �����Դϴ�.
		strRst = strRst & "		<ORIGIN2><![CDATA[]]></ORIGIN2>"									'������ ������ | ������ �� ������ �Է��ϼ���.
		strRst = strRst & "		<EXPIRE_DM><![CDATA[]]></EXPIRE_DM>"								'��ȿ���� | ����8�ڸ� �Է��ϼ��� ��:20100401
		strRst = strRst & "		<SUPPLY_SAVE_YN><![CDATA[N]]></SUPPLY_SAVE_YN>"						'�������ܿ��� | ������ ���� ���θ� Y or N�� �Է��ϼ���. "Y" �Է½� ������ ���� �׸� üũ�˴ϴ�.
		strRst = strRst & "		<DESCRITION><![CDATA[]]></DESCRITION>"								'�����ڸ޸� | ������ �޸� ������ �Է��ϼ���
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetItemRegParameter = strRst
	End Function

	'��ǰ ��� ���� XML
	Public Function getSabangnetSimpleEditItemParameter(ichgSellyn)
		Dim strRst, tmpStatus

		If ichgSellyn <> "" Then
			Select Case ichgSellyn
				Case "Y"	tmpStatus = 2
				Case "N"	tmpStatus = 3
			End Select
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#���� �α��� ���̵�
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#���ݿ��� �߱� ���� ����Ű
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#�������� | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'��ü�ڵ� ��ȯ���� | ��� ������ ����� ��ü�ڵ� ǥ���� (Y : ��ȯ, NULL : ����)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<GOODS_NM><![CDATA["&getItemNameFormat&"]]></GOODS_NM>"				'#��ǰ�� | �ѱ۱��� 50�ڸ����� ��밡���ϸ� , HTML �±� ����� �Ұ��մϴ�.
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#��ü��ǰ�ڵ� | �ڻ翡�� ����ϴ� ��ǰ�ڵ带 �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<STATUS>"&tmpStatus&"</STATUS>"
		strRst = strRst & "		<GOODS_COST><![CDATA["&Clng(GetRaiseValue(FBuycash/10)*10)&"]]></GOODS_COST>"	'#���� | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & "		<GOODS_PRICE><![CDATA["& Clng(MustPrice/10)*10 &"]]></GOODS_PRICE>"	'#�ǸŰ� | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & "		<GOODS_CONSUMER_PRICE><![CDATA["&Clng(FOrgPrice/10)*10&"]]></GOODS_CONSUMER_PRICE>"	'#TAG��(�Һ��ڰ�) | �Է½� ù���ڴ� �ݵ�� ( �� ) ������Ʈ����(ENTER����Key)�� �����ؾ� �ϸ� ���� ���̿� ( , ) �޸��� ���� �ȵ˴ϴ�.
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetSimpleEditItemParameter = strRst
	End Function

	'���θ��� DATA���� XML
	Public Function getSabangnetShoppingMallEditParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<SABANG_GOODS_REGI>"
		strRst = strRst & "	<HEADER>"
		strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#���� �α��� ���̵�
		strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#���ݿ��� �߱� ���� ����Ű
		strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#�������� | YYYYMMDD
		strRst = strRst & "		<SEND_GOODS_CD_RT>Y</SEND_GOODS_CD_RT>"								'��ü�ڵ� ��ȯ���� | ��� ������ ����� ��ü�ڵ� ǥ���� (Y : ��ȯ, NULL : ����)
		strRst = strRst & "	</HEADER>"
		strRst = strRst & "	<DATA>"
		strRst = strRst & "		<MALL_CODE>shop0060</MALL_CODE>"			'#���θ�CODE | ���θ� �ڵ带 �����մϴ�. (���ݸ޴� A>2 ���θ�����(����) �޴� ����)
		strRst = strRst & "		<COMPAYNY_GOODS_CD><![CDATA["& FItemid &"]]></COMPAYNY_GOODS_CD>"	'#��ü��ǰ�ڵ� | �ڻ翡�� ����ϴ� ��ǰ�ڵ带 �����մϴ�. ( 30�ڸ����� )
		strRst = strRst & "		<MALL_PROP1_CD>008</MALL_PROP1_CD>"	'�Ӽ��з��ڵ� | �Ӽ��з��ڵ带 ���� 3�ڸ� �������� �Է��մϴ�. �Ӽ��з��ڵ�� ��ǰ�Ӽ��ڵ� ��ȸ API�� ���� ��ǰ����ȭ���� �Ӽ��з�ǥ�� �����Ͻñ� �ٶ��ϴ�. ��: �Ƿ��� 001�� �Է��մϴ�.
		strRst = strRst & "	</DATA>"
		strRst = strRst & "</SABANG_GOODS_REGI>"
		getSabangnetShoppingMallEditParameter = strRst
	End Function

End Class

Class CSabangnet
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
	Public Sub getSabangnetNotRegOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isnull(c.safetyyn, '') as safetyyn, isnull(c.safetyNum, '') as safetyNum, isnull(c.safetydiv, '') as safetydiv "
		strSql = strSql & "	, isNULL(R.sabangnetStatCD,-9) as sabangnetStatCD "
		strSql = strSql & "	, UC.socname_kor, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_sabangnet_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7')"

		'2020-10-27 ������..������ 1���� �̸��� ����ϰ� �ش޶��..by ����
		' IF (CUPJODLVVALID) then
		' 	strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		' ELSE
		'     strSql = strSql & " and (i.deliveryType<>9)"
	    ' END IF

		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'�ö��/ȭ�����/�ؿ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
'		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.itemdiv <> '06' "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and isnull(R.sabangnetGoodNo, '') = '' "
		strSql = strSql & "		"	& addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FTenCateLarge		= rsget("cate_large")
				FOneItem.FTenCateMid		= rsget("cate_mid")
				FOneItem.FTenCateSmall		= rsget("cate_small")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FItemDiv			= rsget("itemdiv")
				FOneItem.FSmallImage		= rsget("smallImage")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FRegdate			= rsget("regdate")
				FOneItem.FLastUpdate		= rsget("lastUpdate")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FKeywords			= rsget("keywords")
				FOneItem.FVatinclude        = rsget("vatinclude")
				FOneItem.FOrderComment		= db2html(rsget("ordercomment"))
				FOneItem.FOptionCnt			= rsget("optionCnt")
				If isnull(rsget("basicImage600")) or rsget("basicImage600") = "" Then
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				Else
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
				End If
				FOneItem.FMainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FMainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.FIcon1Image		= "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1Image")
				FOneItem.FIcon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.FSourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.FItemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetydiv			= rsget("safetydiv")
				FOneItem.FSabangnetStatCD	= rsget("sabangnetStatCD")
				FOneItem.FDeliverfixday		= rsget("deliverfixday")
				FOneItem.FDeliverytype		= rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
				FOneItem.FBasicimageNm 		= rsget("basicimage")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FMwdiv 			= rsget("mwdiv")
		End If
		rsget.Close
	End Sub

	Public Sub getSabangnetEditOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isnull(c.safetyNum, '') as safetyNum, isNULL(C.safetyDiv, '') as safetyDiv "
		strSql = strSql & "	, m.sabangnetGoodNo, m.sabangnetprice, m.sabangnetSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, m.sabangnetStatCD, UC.socname_kor, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType = 7"
'2020-10-27 ������..������ 1���� �̸��� ����ϰ� �ش޶��..by ����
'		strSql = strSql & "		or ((i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & " 	or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & " 	or i.itemdiv = '06' "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' "
'		strSql = strSql & "		or i.cate_large = '999' "
		strSql = strSql & "		or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.sabangnetStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.sabangnetGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FTenCateLarge		= rsget("cate_large")
				FOneItem.FTenCateMid		= rsget("cate_mid")
				FOneItem.FTenCateSmall		= rsget("cate_small")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FItemDiv			= rsget("itemdiv")
				FOneItem.FSmallImage		= rsget("smallImage")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FRegdate			= rsget("regdate")
				FOneItem.FLastUpdate		= rsget("lastUpdate")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FKeywords			= rsget("keywords")
				FOneItem.FVatinclude        = rsget("vatinclude")
				FOneItem.FOrderComment		= db2html(rsget("ordercomment"))
				FOneItem.FOptionCnt			= rsget("optionCnt")
				If isnull(rsget("basicImage600")) or rsget("basicImage600") = "" Then
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				Else
					FOneItem.FBasicImage		= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage600")
				End If
				FOneItem.FMainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FMainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.FIcon1Image		= "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1Image")
				FOneItem.FIcon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.FSourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.FItemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetydiv			= rsget("safetydiv")
				FOneItem.FSabangnetStatCD	= rsget("sabangnetStatCD")
				FOneItem.FDeliverfixday		= rsget("deliverfixday")
				FOneItem.FDeliverytype		= rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
				FOneItem.FBasicimageNm 		= rsget("basicimage")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FmaySoldOut 		= rsget("maySoldOut")
		End If
		rsget.Close
	End Sub

	Public Sub getSabangnetSimpleEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, m.sabangnetGoodNo, m.sabangnetprice, m.sabangnetSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType = 7"
'2020-10-27 ������..������ 1���� �̸��� ����ϰ� �ش޶��..by ����
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & " 	or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & " 	or i.itemdiv = '06' "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' "
'		strSql = strSql & "		or i.cate_large = '999' "
		strSql = strSql & "		or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_sabangnet_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.sabangnetStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.sabangnetGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSabangnetItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FItemname			= db2html(rsget("itemname"))
				FOneItem.FSabangnetGoodNo	= rsget("sabangnetGoodNo")
				FOneItem.FSabangnetprice	= rsget("sabangnetprice")
				FOneItem.FSabangnetSellYn	= rsget("sabangnetSellYn")
	            FOneItem.FOptionCnt         = rsget("optionCnt")
	            FOneItem.FRegedOptCnt       = rsget("regedOptCnt")
				FOneItem.FOrgPrice			= rsget("orgPrice")
				FOneItem.FOrgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FSellYn			= rsget("sellYn")
				FOneItem.FSaleYn			= rsget("sailyn")
				FOneItem.FIsUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
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
    v = replace(v, "&", "&amp;")
    v = replace(v, """", "&quot;")
	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function
%>