<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "ssg"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST ssgAPIURL = "http://eapi.ssgadm.com"
CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"
CONST CDEFALUT_STOCK = 999
CONST SSGMARGIN = 12									'17%�� ���� �ִ�ġ..12�� ����

Class CSsgItem
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
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FSsgStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FSsgGoodno
	Public FDisplayDate
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public Fcdmkey
	Public Fcddkey
	Public FG9GoodNo
	Public FMapCnt
	Public FMwdiv
	Public FItemsize
	Public FItemsource

	Public FNotinCate
	Public FSafeAuthType
	Public FAuthItemTypeCode
	Public FIsChildrenCate
	Public FOverlap

	Public Function getLimitEa()
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
		getLimitEa = ret
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	Public Function getFiftyUpDown()
		Dim strSql, zoptaddprice, tmpPrice
		If FOptionCnt = 0 Then
			getFiftyUpDown = "N"
		Else
			strSql = ""
			strSql = strSql &" SELECT Max(optaddprice) optaddprice "
			strSql = strSql &" FROM db_item.dbo.tbl_item_option "
			strSql = strSql &" WHERE itemid = '"&FItemid&"' "
			rsget.Open strSql,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				zoptaddprice = rsget("optaddprice")
			End If
			rsget.Close

			If zoptaddprice = 0 Then
				getFiftyUpDown = "N"
			Else
				tmpPrice = Clng(MustPrice / 2)
				If tmpPrice > zoptaddprice Then
					getFiftyUpDown = "N"
				Else
					getFiftyUpDown = "Y"
				End If
			End If
		End If
	End Function

	'// SSG �Ǹſ��� ��ȯ
	Public Function getSsgSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getSsgSellYn = "Y"
			Else
				getSsgSellYn = "N"
			End If
		Else
			getSsgSellYn = "N"
		End If
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

	Public Function getSourcearea()
		Dim arrAreaName, arrAreaCode, i
		arrAreaName = Array("�������", "��ǰ������", "����ĥ��", "�ϼ��δ뼭��", "�ߺϺδ뼭��", "�߼��δ뼭��", "�ߵ��δ뼭��", "�����δ뼭��", "�����δ뼭��", "���غδ뼭��", "�ϴ뼭��", "����븣����", "��������", "����Ƹ���Ƽ��", "��Ŭ����", "���ε���", "������", "�߱�", "�Ϻ�", "�̱���", "��Ʈ��", "��ī��", "��������", "�ϼ��������", "�ߺϺ������", "�߼��������", "�ߵ��������", "�����������", "��Ż����", "���⿡", "����", "������", "����", "�븸", "����", "����", "���̾Ƴ�", "�����", "���׸���", "�׷�����", "�׷�����", "�����������", "���غ������", "���׼�����", "���η��", "������", "�����ͳ�", "�������", "�η��", "�θ�Ű���ļ�", "��ź", "�Ұ�����", "�����", "��糪��", "����ƶ���", "�丣��", "�찣��", "������", "���Ű��ź", "��ũ���̳�", "�긶����", "��������������", "������", "���װ�", "���̼�", "����Ʈ��þ�", "����Ʈ��Ʈ�׷�����", "���Ի�", "������", "�󼼼�������", "�׸���", "��Ÿ", "�߽���", "����", "���-����", "���̺��", "�����", "����������", "��������ī��ȭ��", "�״�����", "����", "�븣����", "��������", "������", "��ī���", "���ѹα�", "����ũ", "���̴�ī", "��Ƽ��", "���ȸ���", "�����", "���̺�����", "��Ʈ���", "���þ�", "���ٳ�", "������", "�縶�Ͼ�", "����θ�ũ", "���ϴ�", "�����ٽ�Ÿ��", "���ٰ���ī��", "����", "����ũ�γ׽þ� ����", "���ɵ��Ͼ�", "������", "��Ͼ�", "�����ƴϾ�", "�����", "�����̽þ�", "����", "�����", "�𸮼Ž�", "��Ÿ��", "�����ũ", "������", "�����", "����", "�̾Ḷ", "�ٴ�����", "�ٺ��̵���", "��Ƽĭ", "���ϸ�", "��۶󵥽�", "����", "�Ҹ�����", "�ַθ�", "����", "������ī", "����������", "������", "������", "������", "���ι�Ű��", "���κ��Ͼ�", "�ø���", "�ÿ��󸮿�", "�ƶ����̸�Ʈ", "�Ƹ��޴Ͼ�", "�Ƹ���Ƽ��", "���̽�����", "����Ƽ", "���Ϸ���", "������������", "�������Ͻ�ź", "�ȵ���", "������", "�Ӱ��", "��Ƽ���ٺδ�", "����Ʈ����", "������Ͼ�", "���⵵��", "����ٵ���", "����", "����", "����Ʈ����", "�µζ�", "�ʸ���", "�밡��", "ī���彺ź", "īŸ��", "ĳ����", "�ɳ�", "�ڸ��", "�ڽ�Ÿ��ī", "��Ʈ���͸�", "�ݷҺ��", "���", "�����ְ�ȭ��", "���", "�����Ʈ", "ũ�ξ�Ƽ��", "Ű���⽺ź", "Ű���ٽ�", "Ű���ν�", "Ÿ��Ű��ź", "ź�ڴϾ�", "�±�", "���", "�밡", "����ũ�޴Ͻ�ź", "Ƣ����", "Ʈ���ϵ�ٵ�", "�ĳ���", "�Ķ����", "��Ű��ź", "��Ǫ�ƴ���Ͼ�", "�ȶ��", "�ȷ���Ÿ��", "���", "��������", "������", "Ǫ�����丮��", "ȣ��", "ȫ��", "��������", "�̵���Ǿ�", "�̶�ũ", "�̶�", "�̽���", "����Ʈ", "�ε�", "�ڸ���ī", "����", "�������", "�߾Ӿ�����ī��ȭ��", "����Ƽ", "���ٺ��", "����", "ü��", "ĥ��", "ī�޷�", "ī��������", "���紺������", "���緯�þ�", "����̱�", "����߽���", "���ص�", "�˶�ī", "�����ε��׽þ�", "������", "����", "��Ű", "�ε��׽þ�", "�ٷ���", "�˹ٴϾ�", "į�����", "�ɶ���", "�̰���", "���߷�", "����", "�����߱�", "�����", "���ε���", "���غ��ε���", "������Ű��ź", "����", "�뼭��", "�ؿܱ�Ÿ", "�Ƹ�������", "��Ÿ", "���ξ���", "�귣�� ��Ʈ����", "��Ÿ", "����", "����Ʈ����", "�ƽþ�", "�̰�����", "���ι�Ű��", "������", "�߱�", "����", "�������", "�̱�", "��Ʈ���", "����ũ", "�;�Ű��", "į�����", "ĥ��", "��ũ���̳�", "�Ƹ���Ƽ��", "��Ż����", "������", "�ϾƸ޸�ī", "����", "��������", "�븣����", "�ɶ���", "��������", "��Ű", "�ݷҺ��", "����ٵ���", "���Ϸ���", "����", "�Ұ�����", "���⿡", "�̾Ḷ", "�����", "�߽���", "��������ī��ȭ��", "Ƣ����", "�±�", "����Ʈ", "�µζ�", "���κ��Ͼ�", "��ī��", "���̴�ī", "�״�����", "����", "ũ�ξ�Ƽ��", "�ε�", "����", "���ٰ���ī��", "�縶�Ͼ�", "�����", "�׸���", "������", "���", "�Ϻ�", "�̶�", "�ø���", "ȣ��", "�밡��", "�ѱ�", "�����", "��Ʈ��", "���þ�", "������", "��Ű��ź", "�̽���", "������ī", "������", "������", "���׸���", "�ʸ���", "ĳ����", "ü��", "����", "������(Georgia)", "GERMANY", "INDONESIA (R&D GERMANY)", "CHINA", "JAPAN", "Germany(����)", "Italy(��Ż����)", "France(������)", "�̱�/����", "����/������", "����/�̱�", "CHINA OEM", "�߱� OEM", "��Ż���ƽþ�ũ", "���ѹα����������", "ENGLAND", "ITALY", "IYALY", "ȣ��", "MALAYSIA(R&D GERMANY)", "���ѹα� (R&D GERMANY)", "KOREA", "CHINA(R&D FRANCE)", "CHINA (R&D FRANCE)", "�����Ͼ�", "�˹ٴϾ�", "����Ʈ��", "���ɵ��Ͼ�", "fusha", "����", "�ٷ���", "��ī���", "������Ͼ�", "�����ƴϾ�", "���¸�", "��Ÿ", "Ÿ�̿�", "����", "����", "�����", "��Ű�ƿ�����", "Ȳ����", "ȫ��", "������", "�𸮼Ž���ȭ��", "�ڽ�Ÿ��ī", "�Ƹ��޴Ͼ�", "������", "��Ǫ�ƴ����", "�ɳ�", "�����", "�������", "��Ƽ���Ǿ�", "����Ʈ����", "�ε��׽þ�", "�丣��", "��۶󵥽�", "�����̽þ�", "�븸")
		arrAreaCode = Array("1000000235", "2000000033", "1000000217", "1000000218", "1000000219", "1000000220", "1000000221", "1000000222", "1000000223", "1000000224", "1000000225", "1000000226", "1000000227", "1000000228", "1000000229", "1000000230", "1000000001", "1000000002", "1000000003", "1000000004", "1000000005", "1000000199", "1000000201", "1000000202", "1000000203", "1000000204", "1000000205", "1000000206", "1000000006", "1000000007", "1000000008", "1000000009", "1000000010", "1000000011", "1000000012", "1000000013", "1000000014", "1000000015", "1000000016", "1000000017", "1000000018", "1000000207", "1000000208", "1000000075", "1000000076", "1000000077", "1000000079", "1000000080", "1000000081", "1000000082", "1000000083", "1000000084", "1000000085", "1000000086", "1000000087", "1000000132", "1000000133", "1000000134", "1000000135", "1000000136", "1000000088", "1000000089", "1000000090", "1000000091", "1000000092", "1000000093", "1000000094", "1000000999", "1000000672", "1000000000", "1000000019", "1000000056", "1000000057", "1000000058", "1000000021", "1000000022", "1000000023", "1000000024", "1000000025", "1000000026", "1000000027", "1000000028", "1000000029", "1000000030", "1000000031", "1000000032", "1000000033", "1000000034", "1000000035", "1000000036", "1000000037", "1000000038", "1000000039", "1000000040", "1000000041", "1000000042", "1000000043", "1000000044", "1000000045", "1000000048", "1000000049", "1000000050", "1000000051", "1000000052", "1000000053", "1000000020", "1000000047", "1000000046", "1000000054", "1000000055", "1000000059", "1000000060", "1000000061", "1000000062", "1000000063", "1000000064", "1000000066", "1000000067", "1000000068", "1000000070", "1000000071", "1000000072", "1000000073", "1000000074", "1000000096", "1000000097", "1000000098", "1000000100", "1000000101", "1000000102", "1000000103", "1000000104", "1000000105", "1000000106", "1000000107", "1000000108", "1000000110", "1000000111", "1000000112", "1000000113", "1000000114", "1000000115", "1000000116", "1000000117", "1000000118", "1000000120", "1000000121", "1000000122", "1000000123", "1000000124", "1000000125", "1000000126", "1000000128", "1000000129", "1000000130", "1000000131", "1000000195", "1000000196", "1000000156", "1000000157", "1000000159", "1000000160", "1000000161", "1000000162", "1000000163", "1000000164", "1000000165", "1000000166", "1000000167", "1000000168", "1000000169", "1000000170", "1000000171", "1000000172", "1000000173", "1000000174", "1000000175", "1000000177", "1000000178", "1000000179", "1000000181", "1000000182", "1000000183", "1000000184", "1000000185", "1000000186", "1000000187", "1000000188", "1000000189", "1000000190", "1000000191", "1000000192", "1000000197", "1000000198", "1000000137", "1000000138", "1000000139", "1000000140", "1000000141", "1000000142", "1000000143", "1000000145", "1000000146", "1000000147", "1000000148", "1000000149", "1000000150", "1000000151", "1000000152", "1000000153", "1000000154", "1000000155", "1000000209", "1000000210", "1000000211", "1000000212", "1000000213", "1000000214", "1000000215", "1000000099", "1000000127", "1000000176", "1000000144", "1000000069", "1000000119", "1000000158", "1000000194", "1000000109", "1000000180", "1000000193", "1000000216", "1000000234", "1000000231", "1000000232", "1000000233", "2001000013", "2001000021", "2000000079", "2000000043", "2000000003", "2000999999", "2001000024", "1000000990", "1000000259", "2000000048", "2000000044", "2000000041", "2000000038", "2000000035", "2000000060", "2000000030", "2000000029", "2000000023", "2000000014", "2000000009", "1000000200", "2000000063", "2000000062", "2000000051", "2000000042", "2000000056", "2000000037", "2000000028", "2000000011", "2000000007", "2000000006", "2000000076", "2000000072", "2000000068", "2000000065", "2000000046", "2000000045", "2000000047", "2000000031", "2000000027", "2000000024", "2000000022", "2000000020", "2000000004", "2000000069", "2000000067", "2000000055", "2000000049", "2000000039", "2000000018", "2000000010", "2000000005", "2000000075", "2000000066", "2000000057", "2000000021", "2000000017", "2000000016", "2000000013", "2000000002", "2000000074", "2000000071", "2000000059", "2000000053", "2000000040", "2000000081", "2000000080", "2000000078", "2000000032", "2000000026", "2000000015", "2000000073", "2000000070", "2000000054", "2000000034", "2000000036", "2000000012", "2000000001", "2000000077", "2000000064", "2000000061", "2000000052", "3000000001", "1000000236", "1000000237", "1000000238", "1000000239", "1000000240", "1000000241", "1000000242", "1000000243", "1000000244", "1000000245", "1000000246", "1000000247", "1000000248", "1000000249", "1000000250", "1000000251", "1000000252", "1000000253", "1000000254", "1000000255", "1000000256", "1000000257", "1000000258", "2001000028", "2001000029", "2001000023", "2001000018", "2001000004", "2001000001", "2001000009", "2001000010", "2001000020", "2001000032", "2001000033", "2001000031", "2001000030", "2001000003", "2001000027", "2001000026", "2001000025", "2001000022", "2000000082", "2001000014", "2001000002", "2001000007", "2001000016", "2001000015", "2001000008", "2001000006", "2001000012", "2001000011", "2001000005", "2001000019", "2000000058", "2000000050", "2000000025", "2000000019", "2000000008")

		If FSourcearea = "�ѱ�" Then
			getSourcearea = "���ѹα�"
		End If

		For i =0 To Ubound(arrAreaName)
			If Trim(arrAreaName(i)) = Trim(FSourcearea) Then
				getSourcearea = Trim(arrAreaCode(i))
				Exit For
			End If
		Next

		If getSourcearea = "" Then
			getSourcearea = "1000000000"		''�󼼼�������
		End If
	End Function

	Public Function getShopLeadTime()
		Dim CateLargeMid, leadTime
		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "030331", "055070", "055080"
				leadTime = 15
			Case "040010", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "055100", "055110", "055120"
				leadTime = 10
			Case "050045"
				leadTime = 7
			Case "040011", "040121", "045002", "045003", "050010", "050020", "050030", "050040", "055090", "055222"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select
		getShopLeadTime = leadTime
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getSsgContParamToReg()
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style><br>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_ssg.jpg'></p><br>"

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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_ssg.jpg"">")
		getSsgContParamToReg = strRst
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

	Public Function getCertInfoParam(iCode)
		Dim strRst, strSql, isChild, isSafe, isElec, isHarm
		Dim chldCertYn, chldCertDivCd, chldCertNo
		Dim certKind, certYn, certDivCd, certNo
		strSql = ""
		strSql = strSql & " SELECT TOP 1 chldCertTgtYn, safeCertTgtYn, elecCertTgtYn, harmCertTgtYn "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_ssg_mmg_category] "
		strSql = strSql & " WHERE stdCtgDclsId = '"&iCode&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			isChild	= rsget("chldCertTgtYn")
			isSafe	= rsget("safeCertTgtYn")
			isElec	= rsget("elecCertTgtYn")
			isHarm	= rsget("harmCertTgtYn")
		End If
		rsget.Close

		If (FSafetyyn = "Y") And (FsafetyDiv = "50") Then
			chldCertYn		= "Y"
			chldCertDivCd	= "10"
			chldCertNo		= FSafetyNum
		Else
			chldCertYn		= "N"
		End If

		If isChild = "Y" Then certKind = "6000000001"
		If isSafe = "Y" Then certKind = "6000000002"
		If isElec = "Y" Then certKind = "6000000003"
		If isHarm = "Y" Then certKind = "6000000004"

		If (FSafetyyn = "Y") AND (FSafetyNum <> "")  Then
			certYn = "Y"
			certDivCd = "10"
			certNo = FSafetyNum
		Else
			certYn = "N"
		End If

		strRst = ""
		strRst = strRst & "	<chldCert>"
		strRst = strRst & "		<chldCertYn>"&chldCertYn&"</chldCertYn>" 									'#������� ����
		strRst = strRst & "		<chldCertDivCd>"&chldCertDivCd&"</chldCertDivCd>"  							'������� ���� (commCd:I368) | (������� ���ΰ� Y�� ��쿡�� �ʼ�) 10 : �����������, 20 : ����Ȯ�δ��, 30 : ���������ռ�Ȯ��
		strRst = strRst & "		<chldCertNo>"&chldCertNo&"</chldCertNo>" 									'������ȣ | ������� ������ 10, 20 �ϰ�쿡�� �ʼ�-
		strRst = strRst & "	</chldCert>"
		If certKind <> "" Then
			strRst = strRst & "	<certInfos>"
			strRst = strRst & "		<certInfo>"
	'����..�Ʒ� ���� certKin -> 6000000004 �̰ɷ� ������ �ְ� ����ϴ� ��ϵǳ�? ��ǰ�ڵ� : 366690 .. 2017-12-20 19:52 ������
			strRst = strRst & "			<certKind>"&certKind&"</certKind>"										'#�������� (commCd:I387) | ������� ī�װ� �� ��� �ʼ�..6000000001 : ������� ��󿩺�, 6000000002 : �������� ��󿩺�, 6000000003 : �������� ���ռ��� ��󿩺�, 6000000004 : ���ؿ����ǰ ǥ�ô�󿩺�
			strRst = strRst & "			<certYn>"&certYn&"</certYn>"											'#���� ����
			strRst = strRst & "			<certDivCd>"&certDivCd&"</certDivCd>"									'���� ���� (commCd:I368) | �������ΰ� Y�̰� ���������� (certKind=6000000001 | 6000000002) �� ��� �ʼ�..10 : �����������, 20 : ����Ȯ�δ��, 30 : ���������ռ�Ȯ��
			strRst = strRst & "			<certNo>"&certNo&"</certNo>"											'������ȣ | ���� ������ 10, 20 �ϰ�쿡�� �ʼ�-
			strRst = strRst & "		</certInfo>"
			strRst = strRst & "	</certInfos>"
		End If
		getCertInfoParam = strRst
'response.write strRst
'response.end
	End Function

	Public Function getSsgAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		strRst = strRst & "	<itemImgs>"
		strRst = strRst & "		<imgInfo>"
		strRst = strRst & "			<dataSeq>1</dataSeq>"													'#�ڷ����
		strRst = strRst & "			<dataFileNm><![CDATA["&FbasicImage&"]]></dataFileNm>"					'#�ڷ����ϸ�
		strRst = strRst & "			<rplcTextNm>��ǥ�̹���</rplcTextNm>"												'#��ü �ؽ�Ʈ ��
		strRst = strRst & "		</imgInfo>"

		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=2 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "		<imgInfo>"
					strRst = strRst & "			<dataSeq>"&i&"</dataSeq>"													'#�ڷ����
					strRst = strRst & "			<dataFileNm><![CDATA[http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & "]]></dataFileNm>"					'#�ڷ����ϸ�
					strRst = strRst & "			<rplcTextNm>��ǰ �̹���"&i&"</rplcTextNm>"												'#��ü �ؽ�Ʈ ��
					strRst = strRst & "		</imgInfo>"
				End If
				rsget.MoveNext
				If i>=9 Then Exit For
			Next
		End If
		rsget.Close
		strRst = strRst & "	</itemImgs>"
'		strRst = strRst & "	<qualityViewImgs>"
'		strRst = strRst & "		<imgInfo>"
'		strRst = strRst & "			<dataSeq></dataSeq>"													'#�ڷ����
'		strRst = strRst & "			<dataFileNm></dataFileNm>"												'#�ڷ����ϸ�
'		strRst = strRst & "			<rplcTextNm></rplcTextNm>"												'#��ü �ؽ�Ʈ ��
'		strRst = strRst & "		</imgInfo>"
'		strRst = strRst & "	</qualityViewImgs>"
		getSsgAddImageParam = strRst
	End Function

	Public Function getRegedOptionCnt
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as Cnt  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption "
		sqlStr = sqlStr & " WHERE mallid= 'ssg' "
		sqlStr = sqlStr & " and itemoption <> '0000' "
		sqlStr = sqlStr & " and itemid=" & FItemid
		rsget.Open sqlStr,dbget,1
			getRegedOptionCnt = rsget("Cnt")
		rsget.Close
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

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
		i = 0
		If Flimityn = "Y" Then
			sqlStr = ""
			sqlStr = sqlStr & "SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option where itemid = '"&iitemid&"' and optaddprice > 0 "
			rsget.Open sqlStr,dbget,1
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
				rsget.Open sqlStr,dbget,1
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
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getSsgCategoryParam()
		Dim sqlStr, i, standardCode, arrDepthCode, arrSiteNo
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 2 stdCtgDClsCd, depthCode, siteNo "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		sqlStr = sqlStr & " WHERE tenCateLarge = '"& FtenCateLarge &"' "
		sqlStr = sqlStr & " and tenCateMid = '"& FtenCateMid &"' "
		sqlStr = sqlStr & " and tenCateSmall = '"& FtenCateSmall &"' "
		sqlStr = sqlStr & " ORDER BY siteNo DESC "
		rsget.Open sqlStr, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				standardCode		= rsget("stdCtgDClsCd")
				arrDepthCode		= arrDepthCode & rsget("depthCode") & ","
				arrSiteNo			= arrSiteNo & rsget("siteNo") & ","
				rsget.MoveNext
			Next
			arrDepthCode = RightCommaDel(arrDepthCode)
			arrSiteNo = RightCommaDel(arrSiteNo)
		End If
		rsget.Close
		getSsgCategoryParam = standardCode & "|_|" & arrDepthCode & "|_|" & arrSiteNo
	End Function

	Public Function getSsgOptParamtoEDIT()
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, requireDetailStr, i
		Dim itemoption, outmalloptcode, outmalloptName, optlimityn, isusing, optsellyn, opt1name, opt2name, opt3name, preged, optNameDiff, oopt, optaddprice
		Dim itemSellTypeCd, OptTypeNm1, OptTypeNm2, OptTypeNm3, optLimit, arrOptionname
		Dim arrRows, isOptionExists, sellStatCd
		Dim arrOptTypeNm

		If FOptionCnt = 0 Then			'��ǰ
			itemSellTypeCd = "10"
		Else
			itemSellTypeCd = "20"
		End If
		strRst = ""
		strRst2 = ""
		strRst3 = ""
		strRst = strRst & "	<itemSellTypeCd>"&itemSellTypeCd&"</itemSellTypeCd>"							'#��ǰ�Ǹ������ڵ� (commCd:I006) | 10 : �Ϲ�, 20 : �ɼ�
		strRst = strRst & "	<itemSellTypeDtlCd>10</itemSellTypeDtlCd>"

		If (FOptionCnt > 0) Then
			strRst = strRst & "	<uitems>"

			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				Do until rsget.EOF
					arrOptTypeNm = arrOptTypeNm & Replace(db2Html(rsget("optionTypeName")),",","")
					rsget.MoveNext
					If Not(rsget.EOF) Then arrOptTypeNm = arrOptTypeNm & ","
				Loop
			End If
			rsget.Close
			arrOptTypeNm = Split(arrOptTypeNm, ",")

			strSql = "EXEC db_item.dbo.usp_Ten_OutMall_Ssg_optEditParamList '"&CMallName&"'," & FItemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If chkMultiOpt Then					'###################### ���߿ɼ��� �� '######################
				Select Case Ubound(arrOptTypeNm)
					Case "1"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = ""
					Case "2"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = Trim(arrOptTypeNm(2))
				End Select

				For i = 0 To UBound(ArrRows,2)
					itemoption		= ArrRows(1,i)
					outmalloptcode	= ArrRows(2,i)
					outmalloptName	= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
					optlimit		= ArrRows(4,i)
					optlimityn		= ArrRows(5,i)
					isusing			= ArrRows(6,i)
					optsellyn		= ArrRows(7,i)
					opt1name		= ArrRows(8,i)
					opt2name		= ArrRows(9,i)
					opt3name		= ArrRows(10,i)
					preged			= (ArrRows(11,i)=1)
					optNameDiff		= (ArrRows(12,i)=1)
					oopt			= ArrRows(13,i)
					optaddprice		= ArrRows(14,i)

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

					If preged = 0 Then
						If (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					Else
						If (optNameDiff) or (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					End If

					strRst = strRst & "		<uitem>"
				If preged = 0 Then
					strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#��ǰID (�ӽù�ȣ)
					strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#��ǰ �ɼ� ������1
					strRst = strRst & "			<uitemOptnNm1>"&opt1name&"</uitemOptnNm1>"					'#��ǰ �ɼ� ��1
					strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'��ǰ �ɼ� ������2
					strRst = strRst & "			<uitemOptnNm2>"&opt2name&"</uitemOptnNm2>"					'��ǰ �ɼ� ��2
					strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'��ǰ �ɼ� ������3
					strRst = strRst & "			<uitemOptnNm3>"&opt3name&"</uitemOptnNm3>"					'��ǰ �ɼ� ��3
					strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'��ǰ �ɼ� ������4
					strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'��ǰ �ɼ� ��4
					strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'��ǰ �ɼ� ������5
					strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'��ǰ �ɼ� ��5
				Else
					strRst = strRst & "			<uitemId>"&outmalloptcode&"</uitemId>"								'#��ǰID
				End If
					strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"						'�ǸŻ����ڵ� | 20:�Ǹ���, 80:�Ͻ��Ǹ�����, 90:�����Ǹ�����
					strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'��� ����
					strRst = strRst & "			<useYn>Y</useYn>"												'��� ����...Y�� �׳� ������ �ǳ�??
					strRst = strRst & "		</uitem>"

					strRst3 = strRst3 & "		<uitemPrc>"
				If preged = 0 Then
					strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#��ǰID (�ӽù�ȣ)
				Else
					strRst3 = strRst3 & "			<uitemId>"&outmalloptcode&"</uitemId>"							'#��ǰID
				End If
					strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
					strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"		'#���ް�
					strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#�ǸŰ�
					strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#������
					strRst3 = strRst3 & "		</uitemPrc>"
				Next
			Else
				For i = 0 To UBound(ArrRows,2)
					itemoption		= ArrRows(1,i)
					outmalloptcode	= ArrRows(2,i)
					outmalloptName	= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
					optlimit		= ArrRows(4,i)
					optlimityn		= ArrRows(5,i)
					isusing			= ArrRows(6,i)
					optsellyn		= ArrRows(7,i)
					opt1name		= ArrRows(13,i)
					opt2name		= ""
					opt3name		= ""
					preged			= (ArrRows(11,i)=1)
					optNameDiff		= (ArrRows(12,i)=1)
					oopt			= ArrRows(13,i)
					optaddprice		= ArrRows(14,i)
					OptTypeNm1		= ArrRows(15,i)

					If OptTypeNm1 = "" Then
						OptTypeNm1 = "����"
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

					If preged = 0 Then
						If (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					Else
						If (optNameDiff) or (isUsing="N") or (optsellyn="N") or (FLimityn = "Y" AND optLimit <= 0) Then
							sellStatCd = "80"
					    Else
							sellStatCd = "20"
						End If
					End If
					strRst = strRst & "		<uitem>"
				If preged = 0 Then
					strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#��ǰID (�ӽù�ȣ)
					strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#��ǰ �ɼ� ������1
					strRst = strRst & "			<uitemOptnNm1>"&opt1name&"</uitemOptnNm1>"					'#��ǰ �ɼ� ��1
					strRst = strRst & "			<uitemOptnTypeNm2></uitemOptnTypeNm2>"							'��ǰ �ɼ� ������2
					strRst = strRst & "			<uitemOptnNm2></uitemOptnNm2>"									'��ǰ �ɼ� ��2
					strRst = strRst & "			<uitemOptnTypeNm3></uitemOptnTypeNm3>"							'��ǰ �ɼ� ������3
					strRst = strRst & "			<uitemOptnNm3></uitemOptnNm3>"									'��ǰ �ɼ� ��3
					strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'��ǰ �ɼ� ������4
					strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'��ǰ �ɼ� ��4
					strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'��ǰ �ɼ� ������5
					strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'��ǰ �ɼ� ��5
				Else
					strRst = strRst & "			<uitemId>"&outmalloptcode&"</uitemId>"								'#��ǰID
				End If
					strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"						'�ǸŻ����ڵ� | 20:�Ǹ���, 80:�Ͻ��Ǹ�����, 90:�����Ǹ�����
					strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'��� ����
					strRst = strRst & "			<useYn>Y</useYn>"												'��� ����...Y�� �׳� ������ �ǳ�??
					strRst = strRst & "		</uitem>"

					strRst3 = strRst3 & "		<uitemPrc>"
				If preged = 0 Then
					strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#��ǰID (�ӽù�ȣ)
				Else
					strRst3 = strRst3 & "			<uitemId>"&outmalloptcode&"</uitemId>"							'#��ǰID
				End If
					strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
					strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"		'#���ް�
					strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#�ǸŰ�
					strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#������
					strRst3 = strRst3 & "		</uitemPrc>"
				Next
			End If
		strRst = strRst & "	</uitems>"
		End If

		If FitemDiv = "06" Then
			requireDetailStr = ""
			requireDetailStr = requireDetailStr & "	<itemOrdOptns>"
			requireDetailStr = requireDetailStr & "		<itemOrdOptn>"
			requireDetailStr = requireDetailStr & "			<addOrdOptnSeq>1</addOrdOptnSeq>"						'#�߰� �ֹ� �ɼ� ����
			requireDetailStr = requireDetailStr & "			<addOrdOptnNm>�ֹ����۹���</addOrdOptnNm>"				'#�߰� �ֹ� �ɼǸ�
			requireDetailStr = requireDetailStr & "		</itemOrdOptn>"
			requireDetailStr = requireDetailStr & "	</itemOrdOptns>"
		End If

		If FOptionCnt > 0 Then
			strRst2 = strRst2 & "	<uitemPluralPrcs>"
			strRst2 = strRst2 & strRst3
			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// �̸�Ʈ�� �߰�
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
'response.write strRst & requireDetailStr & strRst2
'response.end
		getSsgOptParamtoEDIT = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgOptParamtoREG()
		Dim strRst, strRst2, strRst3, strSql, chkMultiOpt, arrOptTypeNm, requireDetailStr
		Dim itemSellTypeCd, OptTypeNm1, OptTypeNm2, OptTypeNm3, optLimit, itemoption, arrOptionname, optionname1, optionname2, optionname3, optaddprice

		If FOptionCnt = 0 Then			'��ǰ
			itemSellTypeCd = "10"
		Else
			itemSellTypeCd = "20"
		End If
		strRst = ""
		strRst2 = ""
		strRst3 = ""
		strRst = strRst & "	<itemSellTypeCd>"&itemSellTypeCd&"</itemSellTypeCd>"							'#��ǰ�Ǹ������ڵ� (commCd:I006) | 10 : �Ϲ�, 20 : �ɼ�
		strRst = strRst & "	<itemSellTypeDtlCd>10</itemSellTypeDtlCd>"										'#��ǰ�Ǹ��������ڵ� (commCd:I007) | 10 : �Ϲ�, 30 : 30 ��ȹ (�ż������ ��ȹ��ǰ �Ұ���)

		If FOptionCnt > 0 Then
			strRst = strRst & "	<uitems>"
			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				Do until rsget.EOF
					arrOptTypeNm = arrOptTypeNm & Replace(db2Html(rsget("optionTypeName")),",","")
					rsget.MoveNext
					If Not(rsget.EOF) Then arrOptTypeNm = arrOptTypeNm & ","
				Loop
			End If
			rsget.Close
			arrOptTypeNm = Split(arrOptTypeNm, ",")

			If chkMultiOpt Then					'###################### ���߿ɼ��� �� '######################
				Select Case Ubound(arrOptTypeNm)
					Case "1"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = ""
					Case "2"
						OptTypeNm1 = Trim(arrOptTypeNm(0))
						OptTypeNm2 = Trim(arrOptTypeNm(1))
						OptTypeNm3 = Trim(arrOptTypeNm(2))
				End Select

				strSql = ""
				strSql = strSql & " SELECT itemid, itemoption, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optionTypeName, optionname, optaddprice, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE isusing = 'Y' and itemid=" & FItemid &"  "
				strSql = strSql & " ORDER BY itemoption ASC "
				rsget.Open strSql,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

						itemoption = rsget("itemoption")
						arrOptionname = rsget("optionname")
						arrOptionname = Split(arrOptionname, ",")
						optaddprice = rsget("optaddprice")

						Select Case Ubound(arrOptTypeNm)
							Case "1"
								optionname1 = Trim(arrOptionname(0))
								optionname2 = Trim(arrOptionname(1))
								optionname3 = ""
							Case "2"
								optionname1 = Trim(arrOptionname(0))
								optionname2 = Trim(arrOptionname(1))
								optionname3 = Trim(arrOptionname(2))
						End Select
						strRst = strRst & "		<uitem>"
						strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#��ǰID (�ӽù�ȣ)
						strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#��ǰ �ɼ� ������1
						strRst = strRst & "			<uitemOptnNm1>"&optionname1&"</uitemOptnNm1>"					'#��ǰ �ɼ� ��1
						strRst = strRst & "			<uitemOptnTypeNm2>"&OptTypeNm2&"</uitemOptnTypeNm2>"			'��ǰ �ɼ� ������2
						strRst = strRst & "			<uitemOptnNm2>"&optionname2&"</uitemOptnNm2>"					'��ǰ �ɼ� ��2
						strRst = strRst & "			<uitemOptnTypeNm3>"&OptTypeNm3&"</uitemOptnTypeNm3>"			'��ǰ �ɼ� ������3
						strRst = strRst & "			<uitemOptnNm3>"&optionname3&"</uitemOptnNm3>"					'��ǰ �ɼ� ��3
						strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'��ǰ �ɼ� ������4
						strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'��ǰ �ɼ� ��4
						strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'��ǰ �ɼ� ������5
						strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'��ǰ �ɼ� ��5
						strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'��� ����
						strRst = strRst & "			<useYn>Y</useYn>"												'��� ����...Y�� �׳� ������ �ǳ�??
						strRst = strRst & "		</uitem>"

						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#��ǰID (�ӽù�ȣ)
						strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
						strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"											'#���ް�
						strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#�ǸŰ�
						strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#������
						strRst3 = strRst3 & "		</uitemPrc>"
						rsget.MoveNext
					Loop
				End If
				rsget.Close
			Else								'###################### ���Ͽɼ��� �� '######################
				strSql = ""
				strSql = strSql & " SELECT itemid, itemoption, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, isnull(optionTypeName, '') as optionTypeName, optionname, optaddprice, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE isusing = 'Y' and itemid=" & FItemid &"  "
				strSql = strSql & " ORDER BY itemoption ASC "
				rsget.Open strSql,dbget,1
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

						itemoption = rsget("itemoption")
						optionname1 = rsget("optionname")
						OptTypeNm1 = rsget("optionTypeName")
						optaddprice = rsget("optaddprice")
						If OptTypeNm1 = "" Then
							OptTypeNm1 = "����"
						End If
						strRst = strRst & "		<uitem>"
						strRst = strRst & "			<tempUitemId>"&itemoption&"</tempUitemId>"						'#��ǰID (�ӽù�ȣ)
						strRst = strRst & "			<uitemOptnTypeNm1>"&OptTypeNm1&"</uitemOptnTypeNm1>"			'#��ǰ �ɼ� ������1
						strRst = strRst & "			<uitemOptnNm1>"&optionname1&"</uitemOptnNm1>"					'#��ǰ �ɼ� ��1
						strRst = strRst & "			<uitemOptnTypeNm2></uitemOptnTypeNm2>"							'��ǰ �ɼ� ������2
						strRst = strRst & "			<uitemOptnNm2></uitemOptnNm2>"									'��ǰ �ɼ� ��2
						strRst = strRst & "			<uitemOptnTypeNm3></uitemOptnTypeNm3>"							'��ǰ �ɼ� ������3
						strRst = strRst & "			<uitemOptnNm3></uitemOptnNm3>"									'��ǰ �ɼ� ��3
						strRst = strRst & "			<uitemOptnTypeNm4></uitemOptnTypeNm4>"							'��ǰ �ɼ� ������4
						strRst = strRst & "			<uitemOptnNm4></uitemOptnNm4>"									'��ǰ �ɼ� ��4
						strRst = strRst & "			<uitemOptnTypeNm5></uitemOptnTypeNm5>"							'��ǰ �ɼ� ������5
						strRst = strRst & "			<uitemOptnNm5></uitemOptnNm5>"									'��ǰ �ɼ� ��5
						strRst = strRst & "			<baseInvQty>"&optLimit&"</baseInvQty>"							'��� ����
						strRst = strRst & "			<useYn>Y</useYn>"												'��� ����...Y�� �׳� ������ �ǳ�??
						strRst = strRst & "		</uitem>"

						strRst3 = strRst3 & "		<uitemPrc>"
						strRst3 = strRst3 & "			<tempUitemId>"&itemoption&"</tempUitemId>"					'#��ǰID (�ӽù�ȣ)
						strRst3 = strRst3 & "			<siteNo>6004</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
						strRst3 = strRst3 & "			<splprc>"&(MustPrice + optaddprice) * 0.85&"</splprc>"											'#���ް�
						strRst3 = strRst3 & "			<sellprc>"&MustPrice + optaddprice&"</sellprc>"				'#�ǸŰ�
						strRst3 = strRst3 & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"								'#������
						strRst3 = strRst3 & "		</uitemPrc>"
						rsget.MoveNext
					Loop
				End If
				rsget.Close
			End If
			strRst = strRst & "	</uitems>"
		End If

		If FitemDiv = "06" Then
			requireDetailStr = ""
			requireDetailStr = requireDetailStr & "	<itemOrdOptns>"
			requireDetailStr = requireDetailStr & "		<itemOrdOptn>"
			requireDetailStr = requireDetailStr & "			<addOrdOptnSeq>1</addOrdOptnSeq>"						'#�߰� �ֹ� �ɼ� ����
			requireDetailStr = requireDetailStr & "			<addOrdOptnNm>�ֹ����۹���</addOrdOptnNm>"				'#�߰� �ֹ� �ɼǸ�
			requireDetailStr = requireDetailStr & "		</itemOrdOptn>"
			requireDetailStr = requireDetailStr & "	</itemOrdOptns>"
		End If
		If FOptionCnt > 0 Then
			strRst2 = strRst2 & "	<uitemPluralPrcs>"
			strRst2 = strRst2 & strRst3
			strRst2 = strRst2 & Replace(strRst3, "<siteNo>6004</siteNo>", "<siteNo>6001</siteNo>")					'// �̸�Ʈ�� �߰�
			strRst2 = strRst2 & "	</uitemPluralPrcs>"
		End If
		getSsgOptParamtoREG = strRst & requireDetailStr & strRst2
	End Function

	Public Function getSsgItemInfoCdToReg(iareaCode)
		Dim strSql, buf, lp
		Dim mallinfoCd, infoContent
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCdAdd='00000') AND (F.chkDiv='Y') THEN 'Y' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00000') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='Y') THEN '10' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00001') AND (F.chkDiv='N') THEN '20' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='Y') THEN 'O' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00002') AND (F.chkDiv='N') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.infoCdAdd='00003') THEN 'N' "
		strSql = strSql & " 	 WHEN (M.mallinfoCd='0000000011') THEN '"&iareaCode&"' "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035' "
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemid&"' "
		strSql = strSql & " WHERE M.mallid = 'ssg' and IC.itemid='"&FItemid&"' "
		rsget.Open strSql,dbget,1
		buf = ""
		If not rsget.EOF Then
			buf = buf & "	<itemMngPropClsId>"& rsget("infoETC") &"</itemMngPropClsId>"
			buf = buf & "	<itemMngAttrs>"
			Do until rsget.EOF
				infoContent = rsget("infocontent")
				mallinfocd = rsget("mallinfocd")
				buf = buf & "	<itemMngAttr>"
				buf = buf & "		<itemMngPropId>"&mallinfocd&"</itemMngPropId>"
				buf = buf & "		<itemMngCntt><![CDATA["&infoContent&"]]></itemMngCntt>"
				buf = buf & "	</itemMngAttr>"
				rsget.MoveNext
			Loop
			buf = buf & "	</itemMngAttrs>"
		End If
		rsget.Close
		getSsgItemInfoCdToReg = buf
	End Function

	'SSG ��� XML
	Public Function getSsgItemRegParameter()
		Dim strRst, i, sellStatCd, areaCode, shppItemDivCd, shppRqrmDcnt, shppRqrmDcntChngRsnCntt

		sellStatCd = 20
		'################################ ī�װ� �׸� ȣ�� ########################################
		Dim callCategory , standardCateCode, arrDisplayCateCode, arrSiteNum
		callCategory = getSsgCategoryParam()
		standardCateCode = Split(callCategory, "|_|")(0)
		arrDisplayCateCode = Split(Split(callCategory, "|_|")(1), ",")
		arrSiteNum = Split(Split(callCategory, "|_|")(2), ",")
		'##########################################################################################
		'################################### ������  ȣ�� ##########################################
		areaCode = getSourcearea()
		'##########################################################################################
		'################################### ��۱���  ȣ�� #########################################
		shppRqrmDcnt = getShopLeadTime()
		'##########################################################################################
'		If FItemdiv = "06" OR FItemdiv = "16" Then
'			shppItemDivCd = "05"
'			If FRequireMakeDay < 1 Then
'				shppRqrmDcnt = 7
'			Else
'				shppRqrmDcnt = FRequireMakeDay
'			End If
'			shppRqrmDcntChngRsnCntt = "�ֹ����ۻ�ǰ"
'		Else
'			shppItemDivCd = "01"
'			shppRqrmDcnt = 3
'		End If

		shppItemDivCd = "01"
		If getShopLeadTime > 3 Then
			shppRqrmDcntChngRsnCntt = "�ֹ����ۻ�ǰ"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<insertItem>"
		strRst = strRst & "	<itemNm><![CDATA["&getItemNameFormat&"]]></itemNm>"								'#��ǰ��
		strRst = strRst & "	<mdlNm></mdlNm>"																'�𵨸�
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#�귣��ID | �ٹ�����(2000047517)
		strRst = strRst & "	<stdCtgId>"&standardCateCode&"</stdCtgId>"										'#ǥ��ī�װ�ID
		strRst = strRst & "	<sites>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#����Ʈ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"								'#�Ǹ� ���� �ڵ� | 20 : �Ǹ���, 80 : �Ͻ��Ǹ�����
		strRst = strRst & "		</site>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#����Ʈ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<sellStatCd>"&sellStatCd&"</sellStatCd>"								'#�Ǹ� ���� �ڵ� | 20 : �Ǹ���, 80 : �Ͻ��Ǹ�����
		strRst = strRst & "		</site>"
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'��ǰ ���� ���� | 00 : ��ü����, 10 : B2C����, 20 : B2E����
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E ���� ���� | 10 : ��ü ����, 20 : ���� ����, 30 : ȸ���� ����
		strRst = strRst & "	<b2cAplRngCd>20</b2cAplRngCd>"													'B2C ���� ���� | 10 : ����, 20 : ���� (���� ���޻� ����), 70 : ���� ����
		strRst = strRst & "	<itemChrctDivCd>10</itemChrctDivCd>"											'#��ǰ Ư�� ���� �ڵ� | 10 : �Ϲ�, 40 : �̰��� �ͱݼ�, 50 : ����� ����Ʈ, 60 : ��ǰ��, 70 : ���� ������
		strRst = strRst & "	<itemChrctDtlCd></itemChrctDtlCd>"												'#��ǰ Ư�� �� �ڵ� | ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 50 | 60) �� ��� ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 50) => 10 : �Ϲ�, 50 : ��ǰ��, ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 60) => 60 : �ż��� ���� ��ǰ��, 70 : �ܺ� ���� ��ǰ��, 80 : ����Ʈ ī��, 90 : ������ ����Ʈ ī��
		strRst = strRst & "	<exusItemDivCd>10</exusItemDivCd>"												'#���� ��ǰ ���� �ڵ� | 10 : �Ϲ�, 20 : GIFT(�Ϲ�)
		strRst = strRst & "	<exusItemDtlCd>10</exusItemDtlCd>"												'#���� ��ǰ �� �ڵ� | 10 : �Ϲ�, 20 : Ư����
		strRst = strRst & "	<dispAplRngTypeCd>10</dispAplRngTypeCd>"										'#���� ���� ���� ���� �ڵ� | 10 : ��ü (����� + PC), 30 : ����� (����� ���ý� ��ü�� ���� �Ұ�)
		strRst = strRst & "	<speSalestrNo></speSalestrNo>"													'Ư�� ������ ��ȣ (Ư���� (exusItemDtlCd=20)�� ��� �Է�) | �� Ư�����ڵ� API ����
		strRst = strRst & getSsgItemInfoCdToReg(areaCode)
		strRst = strRst & "	<manufcoNm><![CDATA["&Trim(FMakername)&"]]></manufcoNm>"						'#�������
		strRst = strRst & "	<prodManufCntryId>"&areaCode&"</prodManufCntryId>"								'#���� ���� ���� ID | (���� : ��������ȸAPI(listOrplc API))
		strRst = strRst & "	<dispCtgs>"
	For i = 0 to Ubound(arrDisplayCateCode)
		strRst = strRst & "		<dispCtg>"
		strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�, 6005 : SSG.COM
		strRst = strRst & "			<dispCtgId>"&arrDisplayCateCode(i)&"</dispCtgId>"						'#���� ī�װ� ID
		strRst = strRst & "			<repDispOrdr>"&i+1&"</repDispOrdr>"										'#��ǥ ���� ���� | �������, �ߺ� ������� ����. ����Ʈ�� �ִ� 3������ ���� ����
		strRst = strRst & "		</dispCtg>"
	Next
		strRst = strRst & "	</dispCtgs>"
		strRst = strRst & "	<dispStrtDts>"&Replace(Date(), "-", "")&"</dispStrtDts>"						'#���ý����Ͻ�(YYYYMMDD OR YYYYMMDDHH24MISS)
		strRst = strRst & "	<dispEndDts>29991231</dispEndDts>"												'#���������Ͻ�(YYYYMMDD OR YYYYMMDDHH24MISS)
'		strRst = strRst & "	<spDispCtgs>"																	'-------- MayBe ����ī�װ� �� ��.. --------
'		strRst = strRst & "		<dispCtg>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<dispCtgId></dispCtgId>"												'#���� ī�װ� ID
'		strRst = strRst & "			<repDispOrdr></repDispOrdr>"											'#��ǥ ���� ���� | �������, �ߺ� ������� ����. ����Ʈ�� �ִ� 3������ ���� ����
'		strRst = strRst & "		</dispCtg>"
'		strRst = strRst & "	</spDispCtgs>"
		strRst = strRst & "	<srchPsblYn>Y</srchPsblYn>"														'�˻� ���� ����
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(FKeywords))&"]]></itemSrchwdNm>"	'��ǰ�˻����
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'���� ȸ�� ��� (���� �������� ���� ��� ALL) | 10 : �йи�, 20 : �����, 30 : �ǹ�, 40 : ���, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'#�ּ� 1ȸ �ֹ� ���� ����
		strRst = strRst & "	<maxOnetOrdPsblQty>9999</maxOnetOrdPsblQty>"									'#�ִ� 1ȸ �ֹ� ���� ����
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'#�ִ� 1�� �ֹ� ���� ����
		strRst = strRst & "	<adultItemTypeCd>90</adultItemTypeCd>"											'#���� ��ǰ Ÿ�� �ڵ� (commCd:I408) | 10 : ���� ��ǰ, 20 : �ַ� ��ǰ, 90 : �Ϲ� ��ǰ
		strRst = strRst & "	<hriskItemYn>N</hriskItemYn>"													'#�� ���� ��ǰ ����
		strRst = strRst & "	<nitmAplYn>N</nitmAplYn>" 														'#�� ��ǰ ���� ����
'		strRst = strRst & "	<sellPnts>"																		'-------- MayBe ��������Ʈ �� ��.. --------
'		strRst = strRst & "		<sellPnt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<sellpntNm></sellpntNm>"												'#
'		strRst = strRst & "			<dispStrtDts></dispStrtDts>"											'#���� ���� �Ͻ� (YYYYMMDD)
'		strRst = strRst & "			<dispEndDts></dispEndDts>"												'#���� ���� �Ͻ� (YYYYMMDD)
'		strRst = strRst & "			<useYn></useYn>"														'#��� ����
'		strRst = strRst & "		</sellPnt>"
'		strRst = strRst & "	</sellPnts>"
		strRst = strRst & "	<sellCapaUnitCd></sellCapaUnitCd>"												'�Ǹ� �뷮 ���� �ڵ� (commCd:I159) | 01 ea, 02 cc, 03 g, 04 kg, 05 m, 06 ml, 07 mm, 08 ��, 09 ��, 10 ��
		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'�Ǹ� �� �뷮
		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'�Ǹ� ���� �뷮
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'�Ǹ� ���� ����
		strRst = strRst & "	<buyFrmCd>60</buyFrmCd>"														'#���� ���� �ڵ� (commCd:I002) | 10 : ������, 20 : ������2(�Ǹź�), 40 : Ư������, 60 : ����Ź
		strRst = strRst & "	<txnDivCd>"&CHKIIF(FVatInclude="N","20","10")&"</txnDivCd>"						'#���� ���� �ڵ� (commCd:I005) | 10 : ����, 20 : �鼼, 30 : ����
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'���ݼ������ | 1 : ���ް� �ڵ���� (Default), 2 : �ǸŰ� �ڵ����, 3 : ���� �ڵ����..�� �� ������ SALE_PRC_INFO, B2E_PRC �Ѵ� ���� �޴´�. ���� ��� �Է� �޾Ƶ� ��� ������ �ش� �� ������ ���� �ش� ���� �ڵ����� ����.
		strRst = strRst & "	<salesPrcInfos>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"														'#���ް�
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#�ǸŰ�
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#������
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"														'#���ް�
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#�ǸŰ�
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#������
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "	</salesPrcInfos>"
'		strRst = strRst & "	<b2ePrcAplTgts>"
'		strRst = strRst & "		<b2ePrcAplTgt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<b2eMbrcoId></b2eMbrcoId>"												'#B2Eȸ����ID
'		strRst = strRst & "			<b2eSplprc></b2eSplprc>"												'#B2E ���ް�
'		strRst = strRst & "			<b2eSellprc></b2eSellprc>"												'#B2E �ǸŰ�
'		strRst = strRst & "			<b2eMrgrt></b2eMrgrt>"													'#B2E ������
'		strRst = strRst & "		</b2ePrcAplTgt>"
'		strRst = strRst & "	</b2ePrcAplTgts>"
		strRst = strRst & "	<invMngYn>Y</invMngYn>"															'#��� ���� ����
		strRst = strRst & "	<baseInvQty>"&getLimitEa()&"</baseInvQty>"										'#��� ����
		strRst = strRst & "	<invQtyMarkgYn>Y</invQtyMarkgYn>"												'#��� ���� ǥ�� ����
'		strRst = strRst & "	<rsvSaleInfo>"
'		strRst = strRst & "		<aplStrtDt></aplStrtDt>"													'#�����Ǹ� ������ (YYYYMMDD)
'		strRst = strRst & "		<aplEndDt></aplEndDt>"														'#�����Ǹ� ������ (YYYYMMDD)
'		strRst = strRst & "		<whoutStrtDt></whoutStrtDt>"												'#��� ���� ���� (YYYYMMDD)
'		strRst = strRst & "		<rstctInvQty></rstctInvQty>"												'#���� �Ǹ� ����
'		strRst = strRst & "	</rsvSaleInfo>"
		strRst = strRst & getSsgOptParamtoREG()
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'#��ۻ�ǰ�����ڵ� (commCd:I070) | 01 : �Ϲ�, 02 : �ؿܱ��Ŵ���, 03 : ��ġ(����), 04 : ��ġ(����), 05 : �ֹ�����, 06 : �ؿ������
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'���ⱹ��(�ؿ� ����� ���ⱹ)shppItemDivCd=06(�ؿ������) �� ��� �ʼ� | ������ ��ȸ API ����(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'���� ��� ���� ��ȣ | shppItemDivCd=06(�ؿ������) �� ��� �ʼ� 10 : ���� �Է�, 20 : �ʼ� �Է�, 30 : �Է� ����
		strRst = strRst & "	<retExchPsblYn>Y</retExchPsblYn>"												'#��ǰ ��ȯ ���� ����
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'#��� ��ü �ڵ� (commCd:P017) | 31 : �ڻ�â��, 32 : ��üâ��, 41 : ���¾�ü
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'#��� ��� �ڵ� (commCd:P021) | 10 : �ڻ���, 20 : �ù���, 30 : ����湮, 40 : ���, 50 : �̹��, 60 : �̹߼�
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'#������ ��ۿ���
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'#��� �ҿ� �ϼ�
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'#��� �ҿ� �ϼ� ���� ���� | ��ǰ��۱����� �Ϲ�(01) �̰� ��ۼҿ��ϼ��� 4�� �̻��� ��� �ʼ�
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'��ü ��ǰ ��ȣ
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'#��� ��ۺ� ID
		strRst = strRst & "	<retShppcstId>0000011336</retShppcstId>"										'#��ǰ ��ۺ� ID
		strRst = strRst & "	<whoutAddrId>0000006297</whoutAddrId>"											'#��� �ּ� ID
		strRst = strRst & "	<snbkAddrId>0000006297</snbkAddrId>"											'#��ǰ �ּ� ID
		strRst = strRst & "	<frgShppPsblYn>N</frgShppPsblYn>"												'#�ؿ� ��� ���� ����
		strRst = strRst & "	<itemTotWgt></itemTotWgt>"												  		'��ǰ �� ����
		strRst = strRst & "	<hopeShppDdDivCd></hopeShppDdDivCd>"											'��� �߼��� ���� �ڵ� (commCd:I015) | 10 : 15���̳�, 20 : 15������ 30���̳�, 30 : 30������, 90 : �߼��� �ִ� ��¥ ����
		strRst = strRst & "	<hopeShppDdEndDts></hopeShppDdEndDts>"											'��� �߼��� ���� �Ͻ� (YYYYMMDD) | ����߼��� �����ڵ尡 (hopeShppDdEndDts=90) �ϰ�� �ʼ�
		strRst = strRst & getSsgAddImageParam()
		strRst = strRst & "	<itemDesc><![CDATA["&getSsgContParamToReg()&"]]></itemDesc>"					'#��ǰ �� ����
		strRst = strRst & "	<sizeDesc><![CDATA["&FItemsize&"]]></sizeDesc>"									'������ ����ǥ
		strRst = strRst & "	<purchGuideCntt></purchGuideCntt>"												'���� �ȳ� ����
		strRst = strRst & "	<asMemoCntt></asMemoCntt>"														'AS �޸� ����
'		strRst = strRst & "	<qualityFiles>"
'		strRst = strRst & "		<qualityFile>"
'		strRst = strRst & "			<itemDescDivCd></itemDescDivCd>"										'#ǰ�� ���� ���� ���� �ڵ� (commCd:I045) | 61 ����������, 65 ���ԽŰ�����, 63 KC������, 64 ���������, 65 ���ԽŰ�����, 66 ���������, 6B ��Ÿ
'		strRst = strRst & "			<imgFileNm></imgFileNm>" 												'#�̹��� ���� ��ġ
'		strRst = strRst & "		</qualityFile>"
'		strRst = strRst & "	</qualityFiles>"
		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & "	<giftPsblYn>Y</giftPsblYn>"														'#���� ���� ����
		strRst = strRst & "	<shppMsgId></shppMsgId>"														'��� �޽��� ID
		strRst = strRst & "	<ssgstrSellYn></ssgstrSellYn>"													'#SSG �����(�ϳ�) �Ǹ� ����
		strRst = strRst & "	<vodExtnlPathUrl></vodExtnlPathUrl>"											'������ �ܺ� ��� URL (��� ��ü�� ���Ͽ�)
		strRst = strRst & "	<palimpItemYn>N</palimpItemYn>"													'#���� ���� ��ǰ ����
		strRst = strRst & "	<itemSellWayCd>10</itemSellWayCd>"												'#��ǰ �Ǹ� ��� �ڵ� (commCd:I392) | 10 �Ϲ�, 20 ��Ż, 30 ���� ����, 40 �Һ�,
		strRst = strRst & "	<itemStatTypeCd>10</itemStatTypeCd>"											'#��ǰ ���� ���� �ڵ� (commCd:I393) | 10 ����ǰ, 20 �߰�, 30 ����, 40 ����, 50 ��ǰ, 60 ��ũ��ġ
		strRst = strRst & "	<whinNotiYn>N</whinNotiYn>"														'#�԰� �˸� ����
'    <book>		'å���� �ʵ�� ����..
'    </book>
		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'���� ���� ���� ����
		strRst = strRst & "</insertItem>"
		getSsgItemRegParameter = strRst
	End Function

	'SSG ���� XML
	Public Function getssgItemEditParameter(ichgSellYn)
		Dim strRst, i, sellStatCd, areaCode, shppItemDivCd, shppRqrmDcnt, shppRqrmDcntChngRsnCntt
		If ichgSellYn = "Y" Then
			sellStatCd = 20
		Else
			sellStatCd = 80
		End If
		'################################ ī�װ� �׸� ȣ�� ########################################
		Dim callCategory , standardCateCode, arrDisplayCateCode, arrSiteNum
		callCategory = getSsgCategoryParam()
		standardCateCode = Split(callCategory, "|_|")(0)
		arrDisplayCateCode = Split(Split(callCategory, "|_|")(1), ",")
		arrSiteNum = Split(Split(callCategory, "|_|")(2), ",")
		'##########################################################################################
		'################################### ������  ȣ�� ##########################################
		areaCode = getSourcearea()
		'##########################################################################################
		'################################### ��۱���  ȣ�� #########################################
		shppRqrmDcnt = getShopLeadTime()
		'##########################################################################################
'		If FItemdiv = "06" OR FItemdiv = "16" Then
'			shppItemDivCd = "05"
'			If FRequireMakeDay < 1 Then
'				shppRqrmDcnt = 7
'			Else
'				shppRqrmDcnt = FRequireMakeDay
'			End If
'			shppRqrmDcntChngRsnCntt = "�ֹ����ۻ�ǰ"
'		Else
'			shppItemDivCd = "01"
'			shppRqrmDcnt = 3
'		End If

		shppItemDivCd = "01"
		If getShopLeadTime > 3 Then
			shppRqrmDcntChngRsnCntt = "�ֹ����ۻ�ǰ"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst & "<updateItem>"
		strRst = strRst & "	<itemId>"&FSsgGoodno&"</itemId>"												'��ǰID
		strRst = strRst & "	<itemNm><![CDATA["&getItemNameFormat&"]]></itemNm>"								'��ǰ��
		strRst = strRst & "	<mdlNm></mdlNm>"																'�𵨸�
		strRst = strRst & "	<deleteMdlNmYn></deleteMdlNmYn>"												'�𵨸� ��������(�¶��� ��ǰ�� ��츸 ����)
		strRst = strRst & "	<brandId>2000047517</brandId>"													'#�귣��ID | �ٹ�����(2000047517)
		strRst = strRst & "	<sites>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#����Ʈ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<sellStatCd>20</sellStatCd>"											'#�Ǹ� ���� �ڵ� | 20 : �Ǹ���, 80 : �Ͻ��Ǹ�����
		strRst = strRst & "		</site>"
		strRst = strRst & "		<site>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#����Ʈ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<sellStatCd>20</sellStatCd>"											'#�Ǹ� ���� �ڵ� | 20 : �Ǹ���, 80 : �Ͻ��Ǹ�����
		strRst = strRst & "		</site>"
		strRst = strRst & "	</sites>"
		strRst = strRst & "	<itemAplRngTypeCd></itemAplRngTypeCd>"											'��ǰ ���� ���� | 00 : ��ü����, 10 : B2C����, 20 : B2E����
		strRst = strRst & "	<b2eAplRngCd>10</b2eAplRngCd>"													'B2E ���� ���� | 10 : ��ü ����, 20 : ���� ����, 30 : ȸ���� ����
		strRst = strRst & "	<b2cAplRngCd>20</b2cAplRngCd>"													'B2C ���� ���� | 10 : ����, 20 : ���� (���� ���޻� ����), 70 : ���� ����
		strRst = strRst & "	<itemChrctDivCd>10</itemChrctDivCd>"											'��ǰ Ư�� ���� �ڵ� | 10 : �Ϲ�, 40 : �̰��� �ͱݼ�, 50 : ����� ����Ʈ, 60 : ��ǰ��, 70 : ���� ������
		strRst = strRst & "	<itemChrctDtlCd></itemChrctDtlCd>"												'��ǰ Ư�� �� �ڵ� | ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 50 | 60) �� ��� ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 50) => 10 : �Ϲ�, 50 : ��ǰ��, ��ǰ Ư�� ���� �ڵ�(itemChrctDivCd = 60) => 60 : �ż��� ���� ��ǰ��, 70 : �ܺ� ���� ��ǰ��, 80 : ����Ʈ ī��, 90 : ������ ����Ʈ ī��
		strRst = strRst & "	<exusItemDivCd>10</exusItemDivCd>"												'���� ��ǰ ���� �ڵ� | 10 : �Ϲ�, 20 : GIFT(�Ϲ�)
		strRst = strRst & "	<exusItemDtlCd>10</exusItemDtlCd>"												'���� ��ǰ �� �ڵ� | 10 : �Ϲ�, 20 : Ư����
		strRst = strRst & "	<dispAplRngTypeCd>10</dispAplRngTypeCd>"										'���� ���� ���� ���� �ڵ� | 10 : ��ü (����� + PC), 30 : ����� (����� ���ý� ��ü�� ���� �Ұ�)
		strRst = strRst & "	<speSalestrNo></speSalestrNo>"													'Ư�� ������ ��ȣ (Ư���� (exusItemDtlCd=20)�� ��� �Է�) | �� Ư�����ڵ� API ����
		strRst = strRst & "	<sellStatCd>"&sellStatCd&"</sellStatCd>"										'�Ǹ� ���� �ڵ� | 20 : �Ǹ���, 80 : �Ͻ��Ǹ�����, 90 : �����Ǹ�����
		strRst = strRst & getSsgItemInfoCdToReg(areaCode)
		strRst = strRst & "	<manufcoNm><![CDATA["&Trim(FMakername)&"]]></manufcoNm>"						'�������
		strRst = strRst & "	<prodManufCntryId>"&areaCode&"</prodManufCntryId>"								'���� ���� ���� ID | (���� : ��������ȸAPI(listOrplc API))
		strRst = strRst & "	<dispCtgs>"
	For i = 0 to Ubound(arrDisplayCateCode)
		strRst = strRst & "		<dispCtg>"
		strRst = strRst & "			<siteNo>"&arrSiteNum(i)&"</siteNo>"										'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�, 6005 : SSG.COM
		strRst = strRst & "			<delYn></delYn>"														'���� ����
		strRst = strRst & "			<dispCtgId>"&arrDisplayCateCode(i)&"</dispCtgId>"						'#���� ī�װ� ID
		strRst = strRst & "			<repDispOrdr>"&i+1&"</repDispOrdr>"										'��ǥ ���� ���� | �������, �ߺ� ������� ����. ����Ʈ�� �ִ� 3������ ���� ����
		strRst = strRst & "		</dispCtg>"
	Next
		strRst = strRst & "	</dispCtgs>"
		strRst = strRst & "	<dispStrtDts>"&Replace(Date(), "-", "")&"</dispStrtDts>"						'���ý����Ͻ�(YYYYMMDD OR YYYYMMDDHH24MISS)
		strRst = strRst & "	<dispEndDts>29991231</dispEndDts>"												'���������Ͻ�(YYYYMMDD OR YYYYMMDDHH24MISS)
'		strRst = strRst & "	<spDispCtgs>"																	'-------- MayBe ����ī�װ� �� ��.. --------
'		strRst = strRst & "		<dispCtg>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<delYn></delYn>"														'���� ����
'		strRst = strRst & "			<dispCtgId></dispCtgId>"												'#���� ī�װ� ID
'		strRst = strRst & "			<repDispOrdr></repDispOrdr>"											'��ǥ ���� ���� | �������, �ߺ� ������� ����. ����Ʈ�� �ִ� 3������ ���� ����
'		strRst = strRst & "		</dispCtg>"
'		strRst = strRst & "	</spDispCtgs>"
		strRst = strRst & "	<srchPsblYn>Y</srchPsblYn>"														'�˻� ���� ����
		strRst = strRst & "	<itemSrchwdNm><![CDATA["&RightCommaDel(Trim(FKeywords))&"]]></itemSrchwdNm>"	'��ǰ�˻����
		strRst = strRst & "	<aplMbrGrdCd></aplMbrGrdCd>"													'���� ȸ�� ��� (���� �������� ���� ��� ALL) | 10 : �йи�, 20 : �����, 30 : �ǹ�, 40 : ���, 50 : VIP, 90 : VVIP
		strRst = strRst & "	<minOnetOrdPsblQty>1</minOnetOrdPsblQty>"										'�ּ� 1ȸ �ֹ� ���� ����
		strRst = strRst & "	<maxOnetOrdPsblQty>9999</maxOnetOrdPsblQty>"									'�ִ� 1ȸ �ֹ� ���� ����
		strRst = strRst & "	<max1dyOrdPsblQty>9999</max1dyOrdPsblQty>"										'�ִ� 1�� �ֹ� ���� ����
		strRst = strRst & "	<adultItemTypeCd>90</adultItemTypeCd>"											'#���� ��ǰ Ÿ�� �ڵ� (commCd:I408) | 10 : ���� ��ǰ, 20 : �ַ� ��ǰ, 90 : �Ϲ� ��ǰ
		strRst = strRst & "	<hriskItemYn>N</hriskItemYn>"													'�� ���� ��ǰ ����
		strRst = strRst & "	<nitmAplYn>N</nitmAplYn>" 														'�� ��ǰ ���� ����
'		strRst = strRst & "	<sellPnts>"																		'-------- MayBe ��������Ʈ �� ��.. --------
'		strRst = strRst & "		<sellPnt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<sellpntId></sellpntId>"												'#���� ����Ʈ ID
'		strRst = strRst & "			<sellpntNm></sellpntNm>"												'#���� ����Ʈ ��
'		strRst = strRst & "			<dispStrtDts></dispStrtDts>"											'#���� ���� �Ͻ� (YYYYMMDD)
'		strRst = strRst & "			<dispEndDts></dispEndDts>"												'#���� ���� �Ͻ� (YYYYMMDD)
'		strRst = strRst & "			<useYn></useYn>"														'#��� ����
'		strRst = strRst & "		</sellPnt>"
'		strRst = strRst & "	</sellPnts>"
		strRst = strRst & "	<sellCapaUnitCd></sellCapaUnitCd>"												'�Ǹ� �뷮 ���� �ڵ� (commCd:I159) | 01 ea, 02 cc, 03 g, 04 kg, 05 m, 06 ml, 07 mm, 08 ��, 09 ��, 10 ��
		strRst = strRst & "	<sellTotCapa></sellTotCapa>"													'�Ǹ� �� �뷮
		strRst = strRst & "	<sellUnitCapa></sellUnitCapa>"													'�Ǹ� ���� �뷮
		strRst = strRst & "	<sellUnitQty>0</sellUnitQty>"													'�Ǹ� ���� ����
		strRst = strRst & "	<prcMngMthd>1</prcMngMthd>"														'���ݼ������ | 1 : ���ް� �ڵ���� (Default), 2 : �ǸŰ� �ڵ����, 3 : ���� �ڵ����..�� �� ������ SALE_PRC_INFO, B2E_PRC �Ѵ� ���� �޴´�. ���� ��� �Է� �޾Ƶ� ��� ������ �ش� �� ������ ���� �ش� ���� �ڵ����� ����.
		strRst = strRst & "	<salesPrcInfos>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6004</siteNo>"													'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"									'#���ް�
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#�ǸŰ�
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#������
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "		<uitemPrc>"
		strRst = strRst & "			<siteNo>6001</siteNo>"													'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
		strRst = strRst & "			<splprc>"&MustPrice()*0.85&"</splprc>"									'#���ް�
		strRst = strRst & "			<sellprc>"&MustPrice()&"</sellprc>"										'#�ǸŰ�
		strRst = strRst & "			<mrgrt>"&SSGMARGIN&"</mrgrt>"											'#������
		strRst = strRst & "		</uitemPrc>"
		strRst = strRst & "	</salesPrcInfos>"
'		strRst = strRst & "	<chgSalesPrcInfos>"
'		strRst = strRst & "		<uitemPrc>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ, 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<splprc></splprc>"														'#���ް�
'		strRst = strRst & "			<sellprc></sellprc>"													'#�ǸŰ�
'		strRst = strRst & "			<mrgrt></mrgrt>"														'#������
'		strRst = strRst & "			<aplStrtDts></aplStrtDts>"												'#���� ���� �Ͻ�(YYYYMMDDHH24MISS)
'		strRst = strRst & "		</uitemPrc>"
'		strRst = strRst & "	</chgSalesPrcInfos>"
'		strRst = strRst & "	<returnSalesPrcInfos>"
'		strRst = strRst & "		<uitemPrc>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ, 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<splprc></splprc>"														'#���ް�
'		strRst = strRst & "			<sellprc></sellprc>"													'#�ǸŰ�
'		strRst = strRst & "			<mrgrt></mrgrt>"														'#������
'		strRst = strRst & "			<aplStrtDts></aplStrtDts>"												'#���� ���� �Ͻ�(YYYYMMDDHH24MISS)
'		strRst = strRst & "		</uitemPrc>"
'		strRst = strRst & "	</returnSalesPrcInfos>"
'		strRst = strRst & "	<b2ePrcAplTgts>"
'		strRst = strRst & "		<b2ePrcAplTgt>"
'		strRst = strRst & "			<siteNo></siteNo>"														'#����Ʈ ��ȣ | 6001 : �̸�Ʈ��, 6002 : Ʈ���̴�����, 6003 : �н���, 6004 : �ż����, 6009 : �ż����ȭ����, 6200 : �ż���TV���θ�
'		strRst = strRst & "			<b2eMbrcoId></b2eMbrcoId>"												'#B2Eȸ����ID
'		strRst = strRst & "			<b2eSplprc></b2eSplprc>"												'#B2E ���ް�
'		strRst = strRst & "			<b2eSellprc></b2eSellprc>"												'#B2E �ǸŰ�
'		strRst = strRst & "			<b2eMrgrt></b2eMrgrt>"													'#B2E ������
'		strRst = strRst & "		</b2ePrcAplTgt>"
'		strRst = strRst & "	</b2ePrcAplTgts>"
		strRst = strRst & "	<invMngYn>Y</invMngYn>"															'��� ���� ����
		strRst = strRst & "	<baseInvQty>"&getLimitEa()&"</baseInvQty>"										'��� ����
		strRst = strRst & "	<invQtyMarkgYn>Y</invQtyMarkgYn>"												'��� ���� ǥ�� ����
'		strRst = strRst & "	<rsvSaleInfo>"
'		strRst = strRst & "		<aplStrtDt></aplStrtDt>"													'#�����Ǹ� ������ (YYYYMMDD)
'		strRst = strRst & "		<aplEndDt></aplEndDt>"														'#�����Ǹ� ������ (YYYYMMDD)
'		strRst = strRst & "		<whoutStrtDt></whoutStrtDt>"												'#��� ���� ���� (YYYYMMDD)
'		strRst = strRst & "		<rstctInvQty></rstctInvQty>"												'#���� �Ǹ� ����
'		strRst = strRst & "		<rsvSaleEndTp></rsvSaleEndTp>"												'#���� �Ǹ� ����(Y�� �Է½� �����Ǹ� ���� ����)
'		strRst = strRst & "	</rsvSaleInfo>"
'		If ichgSellYn = "Y" Then	'ǰ���� �ش����� ���� ���� �ɼ� �����ϱ�..
			strRst = strRst & getSsgOptParamtoEDIT()
'		End If
		strRst = strRst & "	<shppItemDivCd>"&shppItemDivCd&"</shppItemDivCd>"								'��ۻ�ǰ�����ڵ� (commCd:I070) | 01 : �Ϲ�, 02 : �ؿܱ��Ŵ���, 03 : ��ġ(����), 04 : ��ġ(����), 05 : �ֹ�����, 06 : �ؿ������
		strRst = strRst & "	<exprtCntryId></exprtCntryId>"													'���ⱹ��(�ؿ� ����� ���ⱹ)shppItemDivCd=06(�ؿ������) �� ��� �ʼ� | ������ ��ȸ API ����(listOrplc API)
		strRst = strRst & "	<pcusMngCd></pcusMngCd>"														'���� ��� ���� ��ȣ | shppItemDivCd=06(�ؿ������) �� ��� �ʼ� 10 : ���� �Է�, 20 : �ʼ� �Է�, 30 : �Է� ����
		strRst = strRst & "	<retExchPsblYn>Y</retExchPsblYn>"												'��ǰ ��ȯ ���� ����
		strRst = strRst & "	<shppMainCd>41</shppMainCd>"													'��� ��ü �ڵ� (commCd:P017) | 31 : �ڻ�â��, 32 : ��üâ��, 41 : ���¾�ü
		strRst = strRst & "	<shppMthdCd>20</shppMthdCd>"													'��� ��� �ڵ� (commCd:P021) | 10 : �ڻ���, 20 : �ù���, 30 : ����湮, 40 : ���, 50 : �̹��, 60 : �̹߼�
		strRst = strRst & "	<mareaShppYn></mareaShppYn>"													'������ ��ۿ���
		strRst = strRst & "	<shppRqrmDcnt>"&shppRqrmDcnt&"</shppRqrmDcnt>"									'��� �ҿ� �ϼ�
		strRst = strRst & "	<shppRqrmDcntChngRsnCntt>"&shppRqrmDcntChngRsnCntt&"</shppRqrmDcntChngRsnCntt>"	'��� �ҿ� �ϼ� ���� ���� | ��ǰ��۱����� �Ϲ�(01) �̰� ��ۼҿ��ϼ��� 4�� �̻��� ��� �ʼ�
		strRst = strRst & "	<splVenItemId>"&FItemid&"</splVenItemId>"										'��ü ��ǰ ��ȣ
		strRst = strRst & "	<whoutShppcstId>0000517199</whoutShppcstId>"									'��� ��ۺ� ID
		strRst = strRst & "	<retShppcstId>0000011336</retShppcstId>"										'��ǰ ��ۺ� ID
		strRst = strRst & "	<whoutAddrId>0000006297</whoutAddrId>"											'��� �ּ� ID
		strRst = strRst & "	<snbkAddrId>0000006297</snbkAddrId>"											'��ǰ �ּ� ID
		strRst = strRst & "	<frgShppPsblYn>N</frgShppPsblYn>"												'�ؿ� ��� ���� ����
		strRst = strRst & "	<itemTotWgt></itemTotWgt>"												  		'��ǰ �� ����
		strRst = strRst & "	<hopeShppDdDivCd></hopeShppDdDivCd>"											'��� �߼��� ���� �ڵ� (commCd:I015) | 10 : 15���̳�, 20 : 15������ 30���̳�, 30 : 30������, 90 : �߼��� �ִ� ��¥ ����
		strRst = strRst & "	<hopeShppDdEndDts></hopeShppDdEndDts>"											'��� �߼��� ���� �Ͻ� (YYYYMMDD) | ����߼��� �����ڵ尡 (hopeShppDdEndDts=90) �ϰ�� �ʼ�
		If isImageChanged Then
			strRst = strRst & getSsgAddImageParam()
		End If
		strRst = strRst & "	<itemDesc><![CDATA["&getSsgContParamToReg()&"]]></itemDesc>"					'��ǰ �� ����
		strRst = strRst & "	<sizeDesc><![CDATA["&FItemsize&"]]></sizeDesc>"									'������ ����ǥ
		strRst = strRst & "	<purchGuideCntt></purchGuideCntt>"												'���� �ȳ� ����
		strRst = strRst & "	<asMemoCntt></asMemoCntt>"														'AS �޸� ����
'		strRst = strRst & "	<qualityFiles>"
'		strRst = strRst & "		<qualityFile>"
'		strRst = strRst & "			<itemDescDivCd></itemDescDivCd>"										'#ǰ�� ���� ���� ���� �ڵ� (commCd:I045) | 61 ����������, 65 ���ԽŰ�����, 63 KC������, 64 ���������, 65 ���ԽŰ�����, 66 ���������, 6B ��Ÿ
'		strRst = strRst & "			<imgFileNm></imgFileNm>" 												'#�̹��� ���� ��ġ
'		strRst = strRst & "		</qualityFile>"
'		strRst = strRst & "	</qualityFiles>"
		strRst = strRst & getCertInfoParam(standardCateCode)
		strRst = strRst & "	<giftPsblYn>Y</giftPsblYn>"														'���� ���� ����
		strRst = strRst & "	<shppMsgId></shppMsgId>"														'��� �޽��� ID
		strRst = strRst & "	<ssgstrSellYn></ssgstrSellYn>"													'SSG �����(�ϳ�) �Ǹ� ����
		strRst = strRst & "	<vodExtnlPathUrl></vodExtnlPathUrl>"											'������ �ܺ� ��� URL (��� ��ü�� ���Ͽ�)
		strRst = strRst & "	<palimpItemYn>N</palimpItemYn>"													'���� ���� ��ǰ ����
		strRst = strRst & "	<itemSellWayCd>10</itemSellWayCd>"												'��ǰ �Ǹ� ��� �ڵ� (commCd:I392) | 10 �Ϲ�, 20 ��Ż, 30 ���� ����, 40 �Һ�,
		strRst = strRst & "	<itemStatTypeCd>10</itemStatTypeCd>"											'��ǰ ���� ���� �ڵ� (commCd:I393) | 10 ����ǰ, 20 �߰�, 30 ����, 40 ����, 50 ��ǰ, 60 ��ũ��ġ
		strRst = strRst & "	<whinNotiYn>N</whinNotiYn>"														'�԰� �˸� ����
'    <book>		'å���� �ʵ�� ����..
'    </book>
		strRst = strRst & "	<giftPackPsblYn>N</giftPackPsblYn>"												'���� ���� ���� ����
		strRst = strRst & "</updateItem>"
		getssgItemEditParameter = strRst
	End Function

End Class

Class CSsg
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
	Public FRectMustSellyn

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
	Public Sub getSsgNotRegOneItem
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
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(R.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_ssg_regItem as R on i.itemid = R.itemid"
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
		strSql = strSql & " and i.sellcash > i.buycash "
		strSql = strSql & " and i.itemdiv not in ('08', '09', '21') "
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
		strSql = strSql & "	and isnull(R.ssgGoodNo, '') = '' "
		strSql = strSql & " and cm.mapCnt is Not Null "
'		strSql = strSql & " and (i.mwdiv='M' or i.mwdiv='W') "		'���� or ��Ź
'		strSql = strSql & " and i.deliveryType = 1 "				'�Ĺ�
'2018-01-29 15:00 ������ �ϴ� �ּ�ó��..
'		strSql = strSql & " and ( ((i.mwdiv='M' or i.mwdiv='W') and (i.deliveryType = 1)) OR (i.makerid in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic')) ) "
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSsgItem
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
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSsgStatCD			= rsget("ssgStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FMapCnt 			= rsget("mapCnt")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FRequireMakeDay 	= rsget("requireMakeDay")
		End If
		rsget.Close
	End Sub

	Public Sub getSsgEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		If FRectMustSellyn <> "Y" Then
	        ''//���� ���ܻ�ǰ
	        addSql = addSql & " and i.itemid not in ("
	        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
	        addSql = addSql & "     where stDt < getdate()"
	        addSql = addSql & "     and edDt > getdate()"
	        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
	        addSql = addSql & "     and linkgbn='donotEdit'"
	        addSql = addSql & " )"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, isNULL(C.safetyNum, '') as safetyNum "
		strSql = strSql & "	, isNULL(m.ssgStatCD,-9) as ssgStatCD, cm.mapCnt, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, UC.socname_kor, isNULL(c.requireMakeDay,0) as requireMakeDay, m.ssgGoodNo "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv = '21' "
'		strSql = strSql & " 	or i.mwdiv not in ('M', 'W') "
'		strSql = strSql & " 	or i.deliveryType <> 1 "
'2018-01-29 15:00 ������ �ϴ� �ּ�ó��..
'		strSql = strSql & "		or ( ((i.mwdiv not in ('M', 'W')) or (i.deliveryType <> 1)) and i.makerid not in ('meaningless01', 'mandarinebrothers', 'fromamour', 'woolly02' ,'dalbampicnic') )"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.itemdiv = '09' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & " 	FROM db_etcmall.dbo.tbl_ssg_cate_mapping "
		strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_ssg_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.ssgGoodNo is Not Null "		'��� ��ǰ��
		'strSql = strSql & " and m.ssgStatCD = 7' "				'���οϷ�� �ֵ鸸 ������ �ȴ���..TEST �غ��� ��
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSsgItem
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
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSsgStatCD			= rsget("ssgStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FMapCnt 			= rsget("mapCnt")
				FOneItem.FMwdiv 			= rsget("mwdiv")
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FRequireMakeDay 	= rsget("requireMakeDay")
				FOneItem.FmaySoldOut		= rsget("maySoldOut")
				FOneItem.FSsgGoodno			= rsget("ssgGoodno")
		End If
		rsget.Close
	End Sub
End Class

'SSG ��ǰ�ڵ� ���
Function getSsgGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ssgGoodNo FROM db_etcmall.dbo.tbl_ssg_regitem WHERE itemid = '"&iitemid&"' "
	rsget.Open strSql, dbget, 1
		getSsgGoodNo = rsget("ssgGoodNo")
	rsget.Close
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
    v = replace(v, "&", "&amp;")
    v = replace(v, """", "&quot;")
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