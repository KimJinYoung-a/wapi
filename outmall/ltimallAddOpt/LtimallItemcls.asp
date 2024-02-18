<%
'' �����å  3���� ���� 2500
CONST CMAXMARGIN = 14.9			'' MaxMagin��.. '(�Ե�iMall 10%)
CONST CMAXLIMITSELL = 5        '' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CMALLNAME = "lotteimall"
CONST CLTIMALLMARGIN = 11       ''���� 11%
CONST CHEADCOPY = "Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����" ''��Ȱ ����ä�� �ٹ�����
CONST CPREFIXITEMNAME ="[�ٹ�����]"
CONST CitemGbnKey ="K1099999" ''��ǰ����Ű ''�ϳ��� ����
CONST CUPJODLVVALID = TRUE   ''��ü ���ǹ�� ��� ���ɿ���

CONST ENTP_CODE = "011799"                                    '' ���»��ڵ�
CONST MD_CODE   = "0168"                                      '' MD_Code
CONST BRAND_CODE   = "1099329"                                '' �Ե��� �޾ƾ���
CONST BRAND_NAME   = "�ٹ�����(10x10)"                        '' �Ե��� �޾ƾ���
CONST MAKECO_CODE  = "9999"                                   '' �Ե��� �޾ƾ���
CONST CDEFALUT_STOCK = 999       '' ������ ���� �⺻ 999 (���� �ƴѰ��)

Class CLotteiMallItem
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
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FLTiMallGoodNo
	Public FLTiMallTmpGoodNo
	Public FLtiMallPrice
	Public FLtiMallSellYn
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FmaySoldOut
	Public FitemGbnKey
	Public FLtiMallStatCD

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum

	Public FRegedOptionname
	Public FRegedItemname
	Public FNewitemname
	Public FItemnameChange
	Public FItemoption
	Public FOptisusing
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public FOptaddprice
	Public FRealSellprice
	Public FNewItemid
	Public FOptionname
	Public FAdultType

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
	End Function

	Public Function isDiffName
		isDiffName = false
		If (Fitemname <> FRegedItemname) OR (FOptionname <> FRegedOptionname) Then
			isDiffName = True
		End If
	End Function

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
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

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	Function getItemOptNameFormat()
		Dim buf
		buf = replace(getRealItemname,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemOptNameFormat = buf
	End Function

	Public Function getLTiMallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn = "Y" and FisUsing = "Y" Then
			If FLimitYn="N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) Then
				getLTiMallSellYn = "Y"
			Else
				getLTiMallSellYn = "N"
			End if
		Else
			getLTiMallSellYn = "N"
		End If
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, p, r, divBound1, divBound2, divBound3, Keyword1, Keyword2, Keyword3, strRst
		If trim(Fkeywords) = "" Then Exit Function
		'2015-05-06 ������ ���� ������ %�� ���� ������
		Fkeywords  = replace(Fkeywords,"%", "��")
		Fkeywords  = replace(Fkeywords,chr(13), "")
		Fkeywords  = replace(Fkeywords,chr(10), "")
		Fkeywords  = replace(Fkeywords,chr(9), " ")

		If Len(Fkeywords) > 50 Then
			arrRst = Split(Fkeywords,",")
			If Ubound(arrRst) = 0 then
				'������ ������ ���
				arrRst2 = split(arrRst(0)," ")
				If Ubound(arrRst2) > 0 then
					arrRst = split(Fkeywords," ")
				'2013-10-22 ������ ����..ex)826121, 826124
				Else
					'������ �����ݷ��� ���
					arrRst2 = split(arrRst(0),";")
					If Ubound(arrRst2) > 0 then
						arrRst = split(Fkeywords,";")
					End If
				End If
			End If

			If Ubound(arrRst) = 2 Then	'2015-06-29 ������ ���� : ex)769674
				Keyword1 = arrRst(0)
				Keyword2 = arrRst(1)
				Keyword3 = arrRst(2)
			Else
				'Ű���� 1
				divBound1 = Cint(Ubound(arrRst)/3)
				For q = 0 to divBound1
					Keyword1 = Keyword1&arrRst(q)&","
				Next
				If Right(keyword1,1) = "," Then
					keyword1 = Left(keyword1,Len(keyword1)-1)
				End If
				'Ű���� 2
				divBound2 = divBound1 + 1
				For p = divBound2 to divBound2 + divBound1
					Keyword2 = Keyword2&arrRst(p)&","
				Next
				If Right(keyword2,1) = "," Then
					keyword2 = Left(keyword2,Len(keyword2)-1)
				End If

				'Ű���� 3
				divBound3 = divBound2 + divBound1
				For r = divBound3 to Ubound(arrRst)
					Keyword3 = Keyword3&arrRst(r)&","
				Next
				If Right(keyword3,1) = "," Then
					keyword3 = Left(keyword3,Len(keyword3)-1)
				End If
			End If

			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Keyword1
			strRst = strRst & "&sch_kwd_2_nm="&Keyword2
			strRst = strRst & "&sch_kwd_3_nm="&Keyword3
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Fkeywords
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

	'// �ٹ����� ��ǰ�ɼ� �˻�
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold > "&CMAXLIMITSELL&")) "
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold > "&CMAXLIMITSELL&")) "
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

	'// ��ǰ���: MD��ǰ�� �� ���� ī�װ� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallCateParamToReg()
		Dim strSql, strRst, i, ogrpCode
		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " and c.dispLrgNm = '�ٹ�����' "
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//�Ϲݸ��� �⺻ ī�װ���..
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'�⺻ ����ī�װ�
'				Else
'				    IF (ogrpCode=rsget("groupCode")) then
'					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'�߰� ����ī�װ�
'					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getLotteiMallCateParamToReg = strRst
'rw strRst
'response.end
	End Function

	Public Function getLotteiMallGoodDLVDtParams()
		dim strRst
		strRst = ""
		If (FtenCateLarge="055") or (FtenCateLarge="040") then ''����/�к긯 15�Ϸ�
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf (FtenCateLarge="080") or (FtenCateLarge="100") then  ''���/���̺� 5��
			strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
			strRst = strRst & "&dlv_dday=5"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''����/��Ȱ> ��/�̺Ҽ��� or �ֹ���� 10�� - ���ƾ���û 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''������ > ��Ÿ ����Ʈ��� ���̽�  10�� - ���ƾ���û 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''Ȩ/���� > �ſ�   - ���񾾿�û 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ������ 			�ֹ�����15�� 045&cdm=002&cds=001
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ƴ��������			�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > ����/��Ȱ > ����/������ǰ > �������� 			�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ���ڼ��� > ���ڽ� 				�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > ����/��Ȱ > ���ڼ��� > �������ڽ� 			               �ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > Ȩ/���� > ��ǰ�ڽ�/�ٱ��� > �������ڽ�			�ֹ�����10�� cdl=050&cdm=060&cds=070
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > ����ä�� > DIY > �����θ���� 				�ֹ�����10�� 110&cdm=090&cds=040
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''����/��Ȱ > �����μ���  - ���񾾿�û 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
'		ElseIf (FtenCateLarge="025")  then  ''������ 10�� - ���񾾿�û 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
'		    strRst = strRst & "&dlv_dday=10"
		ElseIf ((FitemDiv="06") or (FitemDiv="16")) then    ''�ֹ�(��)���ۻ�ǰ
			strRst = strRst & "&dlv_goods_sct_cd=03"
			If (FrequireMakeDay>7) then
				    strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
			ElseIf (FrequireMakeDay<1) then
				    strRst = strRst & "&dlv_dday=7"
			Else
				    strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
			End If
		Else
			strRst = strRst & "&dlv_goods_sct_cd=01" 																						'��ۻ�ǰ����		(*:�Ϲݻ�ǰ)
			strRst = strRst & "&dlv_dday=3" 																								'��۱���			(*:3���̳�)
		End If
		getLotteiMallGoodDLVDtParams = strRst
	End Function

	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�)
	Public function getLotteiMallOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optDanPoomCD
		chkMultiOpt = false
		optYn = "N"

		strRst = strRst & "&item_mgmt_yn=" & optYn				'��ǰ��������(�ɼ�)
		strRst = strRst & "&opt_nm="							'�ɼǸ�
		strRst = strRst & "&item_list="							'�ɼǻ�
		getLotteiMallOptionParamToReg = strRst
	End Function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallAddImageParamToReg()
		Dim strRst, strSQL, i
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getLotteiMallAddImageParamToReg = strRst
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallItemContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 ������ ž �̹��� �߰�
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>")
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & URLEncodeUTF8(Fordercomment) & "<br>"
		End If

		'#�⺻ ��ǰ����
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&nbsp;"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&nbsp"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,"&"," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(13)," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(10)," ")
		oiMall.FOneItem.Fitemcontent = replace(oiMall.FOneItem.Fitemcontent,chr(9)," ")
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
			Case Else
				strRst = strRst & URLEncodeUTF8(oiMall.FOneItem.Fitemcontent & "<br>")
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
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = Server.URLEncode(rsget("textVal"))
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>") & strtextVal & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg""></div>")
			getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
	End Function

	Public Function getLotteiMallItemInfoCdToReg()
		Dim anjunInfo
        ''������������(�ָ���)
        ''2017-01-06 17:15 ����������..���̸����� sft_cert_org_cd �� ���ٰ� ��..�Ʒ� �ּ����� ->�� �ּ� ǥ��
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			If (FsafetyDiv=10) Then											'������������(KC��ũ)
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS���� -> KS����
			Elseif (FsafetyDiv=20) Then										'�����ǰ ��������
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'�����ǰ�������� -> ��������
			Elseif (FsafetyDiv=30) Then										'KPS �������� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'�����ǰ�������� -> ��������
			Elseif (FsafetyDiv=40) Then										'KPS �������� Ȯ�� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=22"					'�����ǰ��������Ȯ�νŰ� -> ����Ȯ��
			Elseif (FsafetyDiv=50) Then										'KPS ��� ��ȣ���� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS���� -> KS����
			Else
				anjunInfo = ""
			End if
			anjunInfo = anjunInfo & "&sft_cert_no="&Server.URLEncode(FsafetyNum)
		End If

		Dim strRst, strSQL
		Dim mallinfoDiv,mallinfoCd,infoContent, mallinfoCdAll, bufTxt
		Dim rsgetmallinfoDiv, newInfodiv

		If Finfodiv = "47" OR Finfodiv = "48" Then
			newInfodiv = "1" + CSTR(Finfodiv)
		Else
			newInfodiv = ""
		End If

		'���ϸ��� ��ó�� �̴� ����
		Dim YM, ConvertYM, SD
		strSQL = ""
		strSQL = strSQL & " SELECT top 1 F.infocontent, IC.safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		If newInfodiv = "" Then
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		Else
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv= '"& newInfodiv &"'  " & vbcrlf
		End If
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " where IC.itemid='"&Fitemid&"' and M.mallinfocd = '10011' " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		If Not(rsget.EOF or rsget.BOF) then
			YM = rsget("infocontent")
			SD = rsget("safetyDiv")
		Else
			YM = "X"
			SD = "X"
		End If
		rsget.Close

		If YM <> "X" Then
		    YM = replace(YM,".","")
		    YM = replace(YM,"/","")
		    YM = replace(YM,"-","")
		    YM = replace(YM," ","")
		    YM = TRIM(YM)

			If isNumeric(Ym) Then
				ConvertYM = Clng(YM)
			Else
				ConvertYM = "X"
			End If
		Else
			ConvertYM = YM
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE " & vbcrlf

		If SD = "10" Then
			'��ó���� ���� ���� ���
			If ConvertYM = "X" Then
				strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf
			'��ó���� ���� �ִ� ���
			Else
				'��ó���� 2012�� 7�� ������ ���
				If ConvertYM < 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN '�ش����' " & vbcrlf	 '(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� Y, ������ KC(10), 201207��)�� ��
				'��ó���� 2012�� 7�� ������ ���
				ElseIf ConvertYM >= 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf '(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� Y, ������ KC(10), 201207��)�� ��
				End If
			End If
		End If
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10063') THEN IC.safetyNum " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10063') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10205') THEN IC.safetyNum " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10205') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10206') THEN 'KC �������� ��'  " & vbcrlf	'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10206') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN '�ش����'  " & vbcrlf		'(�����ڵ尡 KCC�����̰�), (10x10���� ���������ڵ忩�ΰ� N)�� ��
		strSQL = strSQL & " 	 WHEN M.infoCd='00001' THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN M.infoCd='00002' THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='19008') THEN '������' " & vbcrlf				'�ͱݼ��� ������
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='19008') THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='18008') THEN '��ɼ� �ɻ� ��' " & vbcrlf		'ȭ��ǰ�� ��ɼ� ȭ��ǰ ����
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='18008') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='17008') THEN '��ǰ�������� ���� ���ԽŰ�����' " & vbcrlf		'��ǰ�������� ���� ���ԽŰ� ����	20130215���� �߰�
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='17008') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='M' THEN replace(F.infocontent,'.','') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='C' AND F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infoCd='02004' and F.infocontent='' then '�ش����' " & vbcrlf
		strSql = strSql & " 	 WHEN LEN(F.infocontent) < 2 THEN '��ǰ �� ����' " & vbcrlf
		strSQL = strSQL & " 	 ELSE F.infocontent " & vbcrlf
		strSQL = strSQL & " END AS infoContent, L.shortVal " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		If newInfodiv = "" Then
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		Else
			strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv= '"& newInfodiv &"'  " & vbcrlf
		End If
		strSQL = strSQL & " JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on L.mallid = M.mallid and L.linkgbn='infoDiv21Lotte' and L.itemid ='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'lotteimall' AND IC.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY M.infocd ASC"
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		Dim mat_name, mat_percent, mat_place, material

		If Not(rsget.EOF or rsget.BOF) then
			rsgetmallinfoDiv = rsget("mallinfoDiv")
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))						'��ǰǰ���ڵ�
			Do until rsget.EOF
				mallinfoCd = rsget("mallinfoCd")
				infoContent  = rsget("infoContent")
				If isnull(infoContent) Then
					infoContent = ""
				End If
				infoContent  = replace(infoContent,"%", "��")
				infoContent  = replace(infoContent,chr(13), "")
				infoContent  = replace(infoContent,chr(10), "")
				infoContent  = replace(infoContent,chr(9), " ")
				If mallinfoCd="10085" Then
					If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

						bufTxt = "&mmtr_nm="&mat_name														'�ֿ����
						bufTxt = bufTxt&"&cmps_rt="&mat_percent												'�Է�
						bufTxt = bufTxt&"&mmtr_orpl_nm="&mat_place											'���������
					End If
				End If
				mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &infoContent								'��ǰǰ�� �׸�����
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		Dim isTarget, certParams
		certParams = getNewCertParams(rsgetmallinfoDiv)

		'strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt		'���ȹ� ������
		strRst = certParams & mallinfoDiv & mallinfoCdAll & bufTxt		'���ȹ� ����
		getLotteiMallItemInfoCdToReg = strRst
	End Function

	Public Function getNewCertParams(imallinfoDiv)
		Dim strSql, certNum, safetyDiv, isTarget
		Dim tgtSeq, sctCd, certNo, isRegCert
		Dim strRst
		strRst = ""
		Select Case imallinfoDiv
			Case "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "117", "119", "123", "124", "125", "126", "131", "132", "135"
				isTarget = "Y"		'�����������ǰ�� O
			Case Else
				isTarget = "N"		'�����������ǰ�� X
		End Select
		If isTarget = "Y" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 1 isnull(certNum, '') as certNum, isnull(safetyDiv, '') as safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				certNum		= rsget("certNum")
				safetyDiv	= rsget("safetyDiv")
				isRegCert	= "Y"
			Else
				isRegCert 	= "N"
			End If
			rsget.Close

			If isRegCert = "Y" Then
				tgtSeq = 1
				If imallinfoDiv = "119" Then
					tgtSeq = "2"
				End If

				strRst = strRst & "&sft_cert_tgt_seq="&tgtSeq&""
				Select Case safetyDiv
					Case "10", "40", "70"
						strRst = strRst & "&sft_cert_sct_cd=21"
					Case "20", "50", "80"
						strRst = strRst & "&sft_cert_sct_cd=22"
					Case "30", "60", "90"
						strRst = strRst & "&sft_cert_sct_cd=23"
				End Select
				strRst = strRst & "&sft_cert_no=" & certNum
			Else
				Select Case imallinfoDiv
					Case "106", "115", "117", "124", "125", "126", "131", "132", "135"
						tgtSeq = "2"	'����������� or ���� �� �� ����
					Case "107", "108", "109", "110", "111", "112", "113", "114", "123"
						tgtSeq = "1"	'����ǰ
					Case Else
						tgtSeq = "0"	'�����������ǰ�� X
				End Select

				If imallinfoDiv = "119" Then
					tgtSeq = "1"
				End If

				strRst = strRst & "&sft_cert_tgt_seq="&tgtSeq&""
				strRst = strRst & "&sft_cert_sct_cd="
				strRst = strRst & "&sft_cert_no="
			End If
		End If
		getNewCertParams = strRst
	End Function

	'// �Ե����̸� �Ǹſ��� ��ȯ
	Public Function getLotteiMallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getLotteiMallSellYn = "Y"
			Else
				getLotteiMallSellYn = "N"
			End If
		Else
			getLotteiMallSellYn = "N"
		End If
	End Function

	Public Function getLimitLotteEa()
		Dim ret
		ret = FLimitNo - FLimitSold - 5
		If (ret < 1) Then ret = 0
		getLimitLotteEa = ret
	End Function

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (FLimitYn = "Y") Then
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

    Public Function getLotteiMallItemEditParameter2()
		Dim strRst
		strRst = getLotteiMallItemRegParameter(true)
		getLotteiMallItemEditParameter2 = strRst
    End Function

	'// ��ǰ��� �Ķ���� ����
	Public Function getLotteiMallItemRegParameter(isEdit)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		If (isEdit) Then
		   strRst = strRst & "&goods_req_no="&FLTiMallTmpGoodNo
		End If
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)�귣���ڵ�
		strRst = strRst & "&goods_nm=" & Trim(getItemOptNameFormat)							'(*)���û�ǰ��
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'�𵨹�ȣ(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)��������(1.������, 4.Ư��, 3.Ư���Ǹ�)	�Ե������� 2(�Ǹźи���)�� �����Ǿ�����..���̸��� 2�� ���µ�..�׷��� 4�� �����ߴµ�; ''3�ϵ�: ���� Ȯ��
		strRst = strRst & "&sale_shp_cd=10" 												'(*)�Ǹ������ڵ�(10:����)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FRealSellprice/10)*10)			'(*)�ǸŰ�
		strRst = strRst & "&mrgn_rt="&CLTIMALLMARGIN 										'(*)������(7/1�� �ý��� �����ϸ鼭 11�� �ٲ����..)
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)���鼼�ڵ�(1:����, 2:�鼼)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)MD��ǰ�� �� �ش� ����ī�װ�(��ǰ�������� ī�װ� ������ �� ��..2013-07-02 ����ī�װ� ����API�� ����
		If (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"												'(*)����������(�Ե�����ó�� ����) 2013-06-24 ������
			strRst = strRst & "&inv_qty="&getOptionLimitNo()								'������
		Else
			strRst = strRst & "&inv_mgmt_yn=Y" 												'(*)����������(�Ե�����ó�� ����) 2013-06-24 ������
		    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK									'����Ʈ ���� 99��
		End If
		strRst = strRst & getLotteiMallOptionParamToReg()									'�ɼǸ� �� �ɼǻ� :: ��ǰ��ȣ �߰�
		strRst = strRst & "&add_choc_tp_cd_10="												'��¥�������ɼ�
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=�ֹ����ۻ�ǰ"						 		'�Է����ɼ�
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"										'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"										'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		End If

		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)�������(1:��ü���, 3:���͹��, 4:���Ͱ���, 6:e-�������)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)�������忩��
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)��ۼ���(10:�ù� ,11:��������� ,40:������� ,50:DHL ,60:�ؿܿ��� ,70:�Ϲݿ��� ,80:������)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)��ۻ�ǰ���� �� ��۱���
		strRst = strRst & "&imps_rgn_info_val="												'��ۺҰ�����(10:����,������, 21:����, 22:��������, 23:��õ������, 30:����) �������ǰ��:(�ݷ�)���� �����Ͽ� ���� �Ѱ��� �ݷ����� ����
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 		'(*)�����ڳ�������(0:��ü, 19:19���̻�)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		End If
		strRst = strRst & "&corp_dlvp_sn=525713"						 					'(*)��ǰ��(???) (API_TEST���� ����)
		strRst = strRst & "&corp_rls_pl_sn=525712"						 					'(*)�����(???) (API_TEST���� ����)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)������
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)		'(*)������
		strRst = strRst & "&impr_nm="						 								'�Ǹ���(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)��ǥ�̹���URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'�ΰ��̹���URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)�󼼼���
		strRst = strRst & "&md_ntc_2_FCONT="												'MD����
		strRst = strRst & "&brnd_intro_cont=Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����"		'�귣�� ����
'2013-10-10 ������ ����..���ǻ��� ���� ��ǰ���/�������� ������
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&att_mtr_cont=" &URLEncodeUTF8(ForderComment)					'���ǻ���
		strRst = strRst & "&as_cont="														'AS����
		strRst = strRst & "&gft_nm="														'����ǰ��
		strRst = strRst & "&gft_aply_strt_dtime="											'����ǰ�����Ͻ�
		strRst = strRst & "&gft_aply_end_dtime="											'����ǰ�����Ͻ�
		strRst = strRst & "&gft_fcont="														'����ǰ����
		strRst = strRst & "&corp_goods_no=" & FNewItemid									'��ü��ǰ��ȣ
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'�����尡�ɿ���(��ü��۸�Y ,N) ==> �켱�� Y��..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''����
		getLotteiMallItemRegParameter = strRst
	End Function

	Public Function getLotteiMallOptionParamToEdit()
		Dim ret : ret = ""
		ret = ""
	    If (FLimitYn="Y") Then
		    ret = ret & "&inv_mgmt_yn=Y"
		    ret = ret & "&inv_qty="&getOptionLimitNo()
		    ret = ret & "&item_sale_stat_cd=10"
		Else
			ret = ret & "&inv_mgmt_yn=Y"
			ret = ret & "&inv_qty="&CDEFALUT_STOCK
		    ret = ret & "&item_sale_stat_cd=10"
		END IF

		If (getLTiMallSellYn = "Y") and (getOptionLimitNo > 0) Then										 																'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
			ret = ret & "&sale_stat_cd=10"
		Else
			FSellyn = "N"
			ret = ret & "&sale_stat_cd=20"
		End If
		getLotteiMallOptionParamToEdit = ret
	End Function

	'// ��ǰ���� �Ķ���� ����
	Public Function getLotteiMallItemEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)�Ե����̸� ��ǰ��ȣ
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)�귣���ڵ�
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'�𵨹�ȣ(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)��������(1.������, 4.Ư��, 3.Ư���Ǹ�)	�Ե������� 2(�Ǹźи���)�� �����Ǿ�����..���̸��� 2�� ���µ�..�׷��� 4�� �����ߴµ�; ''3�ϵ�: ���� Ȯ��
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)���鼼�ڵ�(1:����, 2:�鼼)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)�ش� ����ī�װ�(MD��ǰ�� �Ķ��Ÿ�� �ѱ�� �� �������� ������..�Ŵ��� MD��ǰ�� �ѱ�� �Ķ��Ÿ�� ����..���������)
		strRst = strRst & getLotteiMallOptionParamToEdit()
		strRst = strRst & "&add_choc_tp_cd_10="												'��¥�������ɼ�
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=�ֹ����ۻ�ǰ"						 		'�Է����ɼ�
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		End If
		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)�������(1:��ü���, 3:���͹��, 4:���Ͱ���, 6:e-�������)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)�������忩��
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)��ۼ���(10:�ù� ,11:��������� ,40:������� ,50:DHL ,60:�ؿܿ��� ,70:�Ϲݿ��� ,80:������)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)��ۻ�ǰ���� �� ��۱���
		strRst = strRst & "&imps_rgn_info_val="													'��ۺҰ�����(10:����,������, 21:����, 22:��������, 23:��õ������, 30:����) �������ǰ��:(�ݷ�)���� �����Ͽ� ���� �Ѱ��� �ݷ����� ����
		strRst = strRst & "&byr_age_lmt_cd="&Chkiif(IsAdultItem() = "Y", "19", "0")&"" 		'(*)�����ڳ�������(0:��ü, 19:19���̻�)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		End If
		strRst = strRst & "&corp_dlvp_sn=525713"						 					'(*)��ǰ��(???) (API_TEST���� ����)
		strRst = strRst & "&corp_rls_pl_sn=525712"						 					'(*)�����(???) (API_TEST���� ����)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)������
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)		'(*)������
		strRst = strRst & "&impr_nm="						 									'�Ǹ���(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)��ǥ�̹���URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'�ΰ��̹���URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)�󼼼���
		strRst = strRst & "&md_ntc_2_FCONT="													'MD����
		strRst = strRst & "&brnd_intro_cont=Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����"		'�귣�� ����
'2013-10-10 ������ ����..���ǻ��� ���� ��ǰ���/�������� ������
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&attd_mtr_cont=" &URLEncodeUTF8(ForderComment)						'���ǻ���
		strRst = strRst & "&as_cont="															'AS����
		strRst = strRst & "&gft_nm="															'����ǰ��
		strRst = strRst & "&gft_aply_strt_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_aply_end_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_fcont="															'����ǰ����
		strRst = strRst & "&corp_goods_no=" & FNewItemid										'��ü��ǰ��ȣ
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'�����尡�ɿ���(��ü��۸�Y ,N) ==> �켱�� Y��..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''����
		'��� ��ȯ
		getLotteiMallItemEditParameter = strRst
	End Function


End Class

Class CLotteiMall
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
	Public FRectIdx

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

	Public Sub getLtimallNotRegOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname  "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey"
		strSql = strSql & "	, isNULL(R.LtiMallStatCD,-9) as LtiMallStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx "
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'�ö��/ȭ�����/�ؿ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and M.idx not in (Select midx From db_etcmall.dbo.tbl_ltimallAddOption_regItem where LtiMallStatCD > 3 ) "			''LtiMallStatCD>=3 ��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		strSql = strSql & " and o.optsellyn = 'Y' "
		strSql = strSql & " and (o.optlimityn = 'N' or ((o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold >="&CMAXLIMITSELL&"))) "
		strSql = strSql & " and isNULL(c.infodiv,'') not in ('','18','20','22')"
		strSql = strSql & addSql																				'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteiMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FitemGbnKey        = rsget("itemGbnKey")
                FOneItem.FLtiMallStatCD     = rsget("LtiMallStatCD")

                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")

                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")
                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
                FOneItem.FOptionname		= rsget("optionname")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	Public Sub getLtimallEditOneItem
		Dim strSql, addSql, i
        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, R.LtiMallGoodNo, R.LtiMallTmpGoodNo, R.LtiMallprice, R.LtiMallSellYn "
		strSql = strSql & "	, R.accFailCNT, R.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & "	, M.optionname as regedOptionname, M.itemname as regedItemname  "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N' or o.itemoption is null "
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '30') "
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','22')"  ''ȭ��ǰ, ��ǰ�� ����
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or (('1'+c.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123')) and isnull(TR.certNum, '') = '') "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & "	FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
		strSql = strSql & "	JOIN db_item.dbo.tbl_item as i on m.itemid = i.itemid "
		strSql = strSql & "	JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid  "
		strSql = strSql & "	JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx  "
		strSql = strSql & " LEFT JOIN ("
		strSql = strSql & " 	SELECT TOP 1 itemid, isnull(certNum, '') as certNum"
		strSql = strSql & " 	FROM db_item.dbo.tbl_safetycert_tenReg "
		strSql = strSql & " 	WHERE isnull(certNum, '') <> ''"
		strSql = strSql & " ) as TR on i.itemid = TR.itemid"
		strSql = strSql & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid  "
		strSql = strSql & "	LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small  "
		strSql = strSql & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemid = o.itemid and M.itemoption = o.itemoption  "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(R.LtiMallTmpGoodNo, R.LtiMallGoodNo) is Not Null "									'#��� ��ǰ��

'		strSql = ""
'		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname "
'		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
'		strSql = strSql & "	, R.LtiMallGoodNo, R.LtiMallTmpGoodNo, R.LtiMallprice, R.LtiMallSellYn "
'		strSql = strSql & "	, R.accFailCNT, R.lastErrStr "
'		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
'		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
'		strSql = strSql & "	, M.optionname as regedOptionname, M.itemname as regedItemname  "
'        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
'		strSql = strSql & "		or i.isExtUsing='N'"
'		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
'		strSql = strSql & "		or i.sellyn<>'Y'"
'		strSql = strSql & "		or i.deliveryType = 7"
'		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
'		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or isNULL(c.infodiv,'') in ('','18','20','22')"  ''ȭ��ǰ, ��ǰ�� ����
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
'		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
'		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
'		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
'		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
'		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'lotteimall' and M.idx = '"&FRectIdx&"' "
'		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ltimallAddOption_regItem as R on R.midx = M.idx "
'		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
'		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
'		strSql = strSql & " WHERE 1 = 1"
'		strSql = strSql & addSql
'		strSql = strSql & " and isNULL(R.LtiMallTmpGoodNo, R.LtiMallGoodNo) is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteiMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FLTiMallGoodNo		= rsget("LtiMallGoodNo")
				FOneItem.FLtiMallTmpGoodNo	= rsget("LtiMallTmpGoodNo")
				FOneItem.FLtiMallPrice		= rsget("LtiMallprice")
				FOneItem.FLtiMallSellYn		= rsget("LtiMallSellYn")
				FOneItem.Fvatinclude        = rsget("vatinclude")
                FOneItem.FoptionCnt         = rsget("optionCnt")
                FOneItem.FaccFailCNT        = rsget("accFailCNT")
                FOneItem.FlastErrStr        = rsget("lastErrStr")
                ''FOneItem.Fcorp_dlvp_sn      = rsget("returnCode")
                FOneItem.Fdeliverytype      = rsget("deliverytype")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

                FOneItem.FinfoDiv       = rsget("infoDiv")
                FOneItem.Fsafetyyn      = rsget("safetyyn")
                FOneItem.FsafetyDiv     = rsget("safetyDiv")
                FOneItem.FsafetyNum     = rsget("safetyNum")
                FOneItem.FmaySoldOut    = rsget("maySoldOut")

                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")
                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                If not isnull(rsget("itemoption")) Then
                	FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
            	End If
                FOneItem.FOptionname		= rsget("optionname")
	            FOneItem.FRegedOptionname	= rsget("regedOptionname")
	            FOneItem.FRegedItemname		= rsget("regedItemname")
				FOneItem.FAdultType 		= rsget("adulttype")
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

Function getLtimallGoodno(iidx)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 LtimallGoodNo FROM db_etcmall.[dbo].[tbl_ltimallAddOption_regItem] WHERE midx = '"&iidx&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getLtimallGoodno = rsget("ltimallGoodno")
	rsget.Close
End Function
%>