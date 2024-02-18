<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''�ʴ���
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 140
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/googleFeed/") + "\"
Dim FileName: FileName = "googleFeed_temp.xml"
Dim newFileName: newFileName = "googleFeed.xml"
Dim fso, tFile

Function WriteMakeGooleFeedFile(tFile, arrList, byref iLastItemid)
    Dim intLoop, iRow, strSql
    Dim bufstr, isMake
    Dim itemid, deliv, lp, barcode, ArrCateNM
    Dim itemname, designerComment, description, deliveryFixday, adultType
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
		itemid			= arrList(1,intLoop)
		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"������","")
		itemname		= Replace(itemname,"���� ���","")
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		itemname		= Replace(itemname,"""","")

		designerComment	= arrList(3,intLoop)
		deliv 			= arrList(13,intLoop)  ''��ۺ� /2000, 2500, 0

		If designerComment <> "" Then
			description = "��Ȱ����ä�� �ٹ�����- "&Replace(Trim(designerComment),"""","")
		Else
			description = "��Ȱ����ä�� 10x10(�ٹ�����)�� �����μ�ǰ, ���̵���ǰ, ��Ư�� ���׸��� �� �м� ��ǰ ������ ������ ��ſ� ������ �ִ� ���������� ���θ� �Դϴ�."
		End If

		If isNULL(arrList(9,intLoop)) Then
		    ArrCateNM		= ""
		Else
    		ArrCateNM		= Split(arrList(9,intLoop),"||")(0)
			ArrCateNM		= Replace(ArrCateNM, ",", " &gt; " )
        End If

		adultType = arrList(10,intLoop)
		If (adultType="1" or adultType="2") Then
			adultType = "yes"
		Else
			adultType = "no"
		End If

		barcode = "10" & CHKIIF(itemid >= 1000000, Format00(8, itemid), Format00(6, itemid)) & "0000"

		bufstr = "		<item>"
		'** �⺻ ��ǰ ������ **
		bufstr = bufstr & "		<g:id>"&itemid&"</g:id>"						'#[ID] ��ǰ�� ���� �ĺ���
		bufstr = bufstr & "		<g:title><![CDATA["&itemname&"]]></g:title>"	'#[����] ��ǰ �̸�
		bufstr = bufstr & "		<g:description><![CDATA["&description&"]]></g:description>"	'#[����] ��ǰ ���� | ��ǰ�� ��Ȯ�ϰ� �����ϰ� �湮 �������� ����� ��ġ�ϰ� �մϴ�. '���� ���'�� ���� ���θ�� �ؽ�Ʈ, ��� �빮�ڷ� ������ ����, ��Ģ���� �ܱ��� ���ڸ� �����ؼ��� �� �˴ϴ�.
		bufstr = bufstr & "		<g:link>http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"</g:link>"	'#[��ũ] ��ǰ�� �湮 ������
		bufstr = bufstr & "		<g:image_link>http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(5,intLoop)&"</g:image_link>"	'#[�̹���_��ũ] ��ǰ �⺻ �̹����� URL

		strSql = ""
		strSql = strSql & " SELECT TOP 30 gubun, ImgType, addimage_400, addimage_600, addimage_1000 "
		strSql = strSql & " FROM [db_AppWish].[dbo].[tbl_item_addimage] "
		strSql = strSql & " WHERE itemid = '"&itemid&"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open strSql, dbCTget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
			For lp=1 to rsCTget.RecordCount
				If rsCTget("imgType")="0" Then
					bufstr = bufstr & "		<g:additional_image_link>http://webimage.10x10.co.kr/image/add" & rsCTget("gubun") & "/" & GetImageSubFolderByItemid(itemid) & "/" & rsCTget("addimage_400") &"</g:additional_image_link>"	'[�߰�_�̹���_��ũ] ��ǰ�� ���� �߰� �̹����� URL | �ִ�10������
				End If
				rsCTget.MoveNext
				If lp >= 10 Then Exit For
			Next
		END IF
		rsCTget.close

		bufstr = bufstr & "		<g:mobile_link>http://m.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"</g:mobile_link>"	'[�����_��ũ] ����ϰ� ����ũ�� Ʈ���ȿ� ���� URL�� �ٸ� ��� ����Ͽ� ����ȭ�� ��ǰ �湮 ������
		'** ���� �� ��� **
		bufstr = bufstr & "		<g:availability>in stock</g:availability>"										'#[���] ��ǰ ��� | in stock[��� ����], out of stock[��� ����], preorder[���ֹ�]
'		bufstr = bufstr & "		<g:availability_date>YYYY-MM-DD</g:availability_date>"							'[���_����_��¥] ���ֹ� ��ǰ�� ��� ���� ��¥ | preorder[���ֹ�]�� availability[���]�� �����ϴ� ��쿡 �� �Ӽ��� ����մϴ�.
'		bufstr = bufstr & "		<g:cost_of_goods_sold>23000.00 KRW</g:cost_of_goods_sold>"						'[�������] Ư�� ��ǰ �Ǹſ� ������ �������, ������ ȸ�� �������� ���ǵ˴ϴ�. �̷��� ��뿡�� ���, �ӱ�, ȭ���̳� ��Ÿ ��������� ���Ե� �� �ֽ��ϴ�. ��ǰ�� ��������� �����ϸ� ���� ����� �߻��� �����Ͱ� ���� �Ը�� ���� �ٸ� �����׸��� �ľ��� �� �ֽ��ϴ�.
'		bufstr = bufstr & "		<g:expiration_date>YYYY-MM-DD</g:expiration_date>"								'[������] ��ǰ ǥ�ð� �ߴܵǾ�� �ϴ� ��¥ | ������ 30�� �̳��� ��¥�� ����մϴ�.

		If arrList(7,intLoop) <> "Y" Then		'������ �ƴϸ�
			bufstr = bufstr & "		<g:price>"&arrList(5,intLoop)&" KRW</g:price>"								'#[����] ��ǰ ����
		Else
			bufstr = bufstr & "		<g:price>"&arrList(4,intLoop)&" KRW</g:price>"								'#[����] ��ǰ ����
			bufstr = bufstr & "		<g:sale_price>"&arrList(6,intLoop)&" KRW</g:sale_price>"					'[���ΰ�] ��ǰ ���ΰ�
'			bufstr = bufstr & "		<g:sale_price_effective_date>YYYY-MM-DD</g:sale_price_effective_date>"		'[���ΰ�_����_��] ��ǰ�� sale_price�� ����Ǵ� �Ⱓ
		End If
'		bufstr = bufstr & "		<g:unit_pricing_measure>1.5kg</g:unit_pricing_measure>"						'[�ܰ�_å��_����] ��ǰ �Ǹ� ������ ����ġ �� ũ�� | ����: oz, lb, mg, g, kg �뷮(�̱� ��ġ��): floz, pt, qt, gal �뷮 ����: ml, cl, l, cbm  ����: in, ft, yd, cm, m ����: sqft, sqm ������: ct
'		bufstr = bufstr & "		<g:unit_pricing_base_measure>100g</g:unit_pricing_base_measure>"			'[�ܰ�_å��_����_����] ��ǰ�� ���� å�� ���� ����(��: 100ml�� 100ml ������ ������ ����) | unit_pricing_measure[�ܰ�_å��_����]�� �����ϴ� ��쿡 ���û����Դϴ�.
'		bufstr = bufstr & "		<g:installment>"														'[�Һ�] �Һ� ���� ����� ��������
'		bufstr = bufstr & "			<g:months>6</g:months>"													'[����] �����̸� �����ڰ� �����ؾ� �ϴ� �Һ� Ƚ���Դϴ�.
'		bufstr = bufstr & "			<g:amount>50BRL</g:amount>"												'[�����ξ�] ISO 4217 ǥ���� ����� �ϸ� �����ڰ� �ſ� �����ؾ� �ϴ� �ݾ��Դϴ�.
'		bufstr = bufstr & "		</g:installment>"
'		bufstr = bufstr & "		<g:subscription_cost>"													'[����_���] ���� ��ǰ�� ��� ���� ����� ����� �Բ� �����ϴ� ���� �Ǵ� ���� ������� ��������
'		bufstr = bufstr & "			<g:period>����</g:period>"												'#[�Ⱓ] ���� ���� ������ �Ⱓ����, 'month[��]' �Ǵ� 'year[��]' �����Դϴ�.
'		bufstr = bufstr & "			<g:period_length>12</g:period_length>"									'#[�Ⱓ_����] �����ڰ� �����ؾ� �ϴ� �� �Ǵ� �� ���� ���� �Ⱓ�� ����(����)�Դϴ�.
'		bufstr = bufstr & "			<g:amount>50000 KRW</g:amount>"											'#[�����ξ�]   ISO 4217 ǥ���� ����� �ϸ� �����ڰ� �ſ� �����ؾ� �ϴ� �ݾ��Դϴ�. �� �ݾ��� ǥ���� �� ������ �� �����ϵ��� ���� ����� ���� ��ȭ ������ �ݾ��� �ݿø��� �� �ֽ��ϴ�. ������ ������ ���� ����Ʈ�� ǥ�õǴ� �ݾװ� ��Ȯ�� ��ġ�ؾ� �մϴ�.
'		bufstr = bufstr & "		</g:subscription_cost>"
'		bufstr = bufstr & "		<g:loyalty_points>" 													'[����_����Ʈ] (�Ϻ��� �ش�)��ǰ�� ������ �� ���� �޴� ���� ����Ʈ�� ����
'		bufstr = bufstr & "			<g:name>Program A</g:name>"												'#[����Ʈ_��] ��ǰ���� ȹ���� ����Ʈ
'		bufstr = bufstr & "			<g:points_value>100</g:points_value>"									'[�̸�] �Ϻ��� 12�� �Ǵ� �θ��� 24�ڷ� ������ ���� ����Ʈ ������ �̸�
'		bufstr = bufstr & "			<g:ratio>1.0</g:ratio>"													'[����] ��ȭ�� ��ȯ �� ����Ʈ ����(����)
'		bufstr = bufstr & "		</g:loyalty_points>"
		'** ��ǰ ī�װ� **
		bufstr = bufstr & "		<g:google_product_category>956</g:google_product_category>"					'[Google_��ǰ_ī�װ�] ��ǰ�� ���� Google���� ������ ��ǰ ī�װ� (�ϴ� ���̾�� �Ҳ��� �繫��ǰ>�Ϲ� �繫��ǰ>���� ��ǰ�� ��Ī)
		bufstr = bufstr & "		<g:product_type><![CDATA["&ArrCateNM&"]]></g:product_type>"					'[��ǰ_����] ��ǰ�� ���� ������ ��ǰ ī�װ�
		'** ��ǰ �ĺ��� **
		bufstr = bufstr & "		<g:brand><![CDATA["&arrList(12,intLoop)&"]]></g:brand>"						'#[�귣��] ��� �� ��ǰ�� ��� �ʼ������̸� ��ȭ, ����, ���� �귣��� ����
'		bufstr = bufstr & "		<g:gtin>71919219405200</g:gtin>"	'[GTIN] ������ü�� �Ҵ��� GTIN�� �ִ� ��� �� ��ǰ�� ��� | ���� itemstock ���̺��� barcode�� ��� �� �� ������..�ɼǺ��� ������
		bufstr = bufstr & "		<g:mpn>"&barcode&"</g:mpn>"	'[MPN] �� ��ǰ�� ������ü���� �Ҵ��� GTIN�� ���� ��츸 �ش�
		bufstr = bufstr & "		<g:identifier_exists>yes</g:identifier_exists>"								'[�ĺ���_����] ��ǰ�� ��ǰ ���� �ĺ���(UPI) GTIN, MPN, �귣�尡 �ִ��� ���θ� ����Ϸ��� ����մϴ�.
		'** �� ��ǰ ���� **
		bufstr = bufstr & "		<g:condition>new</g:condition>"												'#[����] | new[�� ��ǰ] ���ο� ��ǰ�̳� �������� ��ǰ, ���� ���� ��, refurbished[���� ��ǰ] ,���������� ���� ���·� ������ ��ǰ, ���� ����, ������ ������ ���� �ְ� �ƴ� ���� ����, used[�߰�ǰ] �̹� ���� ��ǰ, ������ ������ �����Ǿ��ų� ������ ����
		bufstr = bufstr & "		<g:adult>"&adultType&"</g:adult>"											'#[����] ��ǰ�� ���ο� �������� ���Ե� ��� | yes[��] no[�ƴϿ�]
'		bufstr = bufstr & "		<g:multipack>6</g:multipack>"												'[��Ű�� ��ǰ] �Ǹ��ڰ� ������ ��Ű�� ��ǰ�� ���ԵǾ� �ǸŵǴ� ���� ��ǰ�� ����
'		bufstr = bufstr & "		<g:is_bundle>no</g:is_bundle>"												'[����_����] 1���� �ֿ� ��ǰ�� �̸� �����ϴ� ���� ��ǰ���� �Ǹ��ڰ� ������ ���� �׷� ��ǰ���� ���
'		bufstr = bufstr & "		<g:energy_efficiency_class>A+</g:energy_efficiency_class>"					'[������_ȿ��_���] ��ǰ�� ������ ��
'		bufstr = bufstr & "		<g:min_energy_efficiency_class>A+++</g:min_energy_efficiency_class>"		'[�ּ�_������_ȿ��_���] ��ǰ�� ������ ��
'		bufstr = bufstr & "		<g:max_energy_efficiency_class>D</g:max_energy_efficiency_class>"			'[�ִ�_������_ȿ��_���] ��ǰ�� ������ ��
'		bufstr = bufstr & "		<g:age_group>infant</g:age_group>"											'[���ɴ�] ��ǰ�� ��� �α���� | newborn[�Ż���] 3���� ����, infant[����] 3����~12����, toddler[����] 1��~5��, kids[���] 5��~13��, adult[����] �Ϲ������� 10�� �̻�
'		bufstr = bufstr & "		<g:color>Black</g:color>"													'[����] ��ǰ�� ����
'		bufstr = bufstr & "		<g:gender>unisex</g:gender>"												'[����] ��ǰ�� ��� ���� | male[����], female[����], unisex[�������]
'		bufstr = bufstr & "		<g:material>leather</g:material>"											'[����] ��ǰ�� ���� �Ǵ� ����
'		bufstr = bufstr & "		<g:pattern>striped</g:pattern>"												'[����] ��ǰ�� ���� �Ǵ� �׷��� ����Ʈ
'		bufstr = bufstr & "		<g:size>XL</g:size>"														'[ũ��] ��ǰ�� ������
'		bufstr = bufstr & "		<g:size_type>regular</g:size_type>"											'[ũ��_����] �Ƿ� ��ǰ�� �� | regular[�Ϲ�], petite[�ڶ�], plus[�÷���], big and tall[�� ������], maternity[�ӻ��]
'		bufstr = bufstr & "		<g:size_system>US</g:size_system>"											'[������_ü��] ��ǰ�� ���Ǵ� ������ ü���� ���� | US, UK, EU, DE, FR, JP, CN(�߱�), IT, BR, MEX, AU
'		bufstr = bufstr & "		<g:item_group_id>AB12345</g:item_group_id>"									'[��ǰ_�׷�_ID] ���� ����(����)���� �����Ǵ� ��ǰ �׷��� ID
		'** ���� ķ���� �� ��Ÿ ���� **
'		bufstr = bufstr & "		<g:ads_redirect>http://www.example.com/product.html</g:ads_redirect>"		'[ads_���𷺼�] ��ǰ �������� �߰� �Ű������� �����ϴ� �� ���Ǵ� URL�Դϴ�. ����ڴ� link[��ũ] �Ǵ� mobile_link[�����_��ũ]�� ����� ���� �ƴ϶� �� URL�� �̵��մϴ�.
'		bufstr = bufstr & "		<g:custom_label_0>���� ��ǰ</g:custom_label_0>"								'[����_��_0] ���� ķ������ ���� �� ���� �����ϱ� ���� ��ǰ�� �Ҵ��ϴ� ���Դϴ�. | �� �Ӽ��� ���� �� �����Ͽ� ��ǰ�� �ִ� 5������ ���� ���� �����մϴ�. custom_label_0[����_��_0], custom_label_1[����_��_1], custom_label_2[����_��_2], custom_label_3[����_��_3], custom_label_4[����_��_4]
'		bufstr = bufstr & "		<g:promotion_id>ABC123</g:promotion_id>"									'[���θ��_ID] ��ǰ�� �Ǹ��� ���θ�ǿ� ������ �� �ִ� �ĺ����Դϴ�.
		'** ������ **
'		bufstr = bufstr & "		<g:excluded_destination>Shopping Ads</g:excluded_destination>"				'[���ܵǴ�_����] Ư�� ������ ���� ķ���ο� ��ǰ�� �������� �ʵ��� �����ϴ� �� ����� �� �ִ� ���� | Shopping Ads[���� ����], Shopping Actions[���� �۾�], Display Ads[���÷��� ����], Surfaces across Google[���� Google ��ǰ�� ����]
'		bufstr = bufstr & "		<g:included_destination>Shopping Ads</g:included_destination>"				'[���ԵǴ�_����] Ư�� ������ ���� ķ���ο� ��ǰ�� �����ϴ� �� ����� �� �ִ� ���� | Shopping Ads[���� ����], Shopping Actions[���� �۾�], Display Ads[���÷��� ����], Surfaces across Google[���� Google ��ǰ�� ����]
		'** ��� **
		bufstr = bufstr & "		<g:shipping>"															'[���] ��ǰ�� ��ۺ�
		bufstr = bufstr & "			<g:country>KR</g:country>"												'[����] ISO 3166 ���� �ڵ�
'		bufstr = bufstr & "			<g:region>MA</g:region>"												'[����] ��, ����, ���� �����մϴ�. �̱�, ����Ʈ���ϸ���, �Ϻ��� �����˴ϴ�. ���� ���ξ� ���� ISO 3166-1 ���� �ڵ带 �����մϴ�(��: CA, NSW, 03).
		bufstr = bufstr & "			<g:service>�Ϲ� ���</g:service>"										'[����] ���� ��� �Ǵ� ��� �ӵ�
		bufstr = bufstr & "			<g:price>"&deliv&" KRW</g:price>"							'#[����] ���� ��ۺ�(�ʿ��� ��� VAT ����)
		bufstr = bufstr & "		</g:shipping>"
'		bufstr = bufstr & "		<g:shipping_label>�ż� ��ǰ</g:shipping_label>"								'[��۹�_��] �Ǹ��� ���� ���� �������� �ùٸ� ��ۺ� �Ҵ��ϱ� ���� ��ǰ�� �Ҵ��ϴ� ��
'		bufstr = bufstr & "		<g:shipping_weight>3kg</g:shipping_weight>"									'[��۹�_�߷�] ��ۺ� ����ϴ� �� ���Ǵ� ��ǰ�� �߷�
'		bufstr = bufstr & "		<g:shipping_length>20cm</g:shipping_length>"								'[��۹�_����] ���� �߷��� ��ۺ� ����ϴ� �� ���Ǵ� ��ǰ�� ����
'		bufstr = bufstr & "		<g:shipping_width>20cm</g:shipping_width>"									'[��۹�_��] ���� �߷��� ��ۺ� ����ϴ� �� ���Ǵ� ��ǰ�� ��
'		bufstr = bufstr & "		<g:shipping_height>20cm</g:shipping_height>"								'[��۹�_����] ���� �߷��� ��ۺ� ����ϴ� �� ���Ǵ� ��ǰ�� ����
'		bufstr = bufstr & "		<g:transit_time_label>�þ�Ʋ ���</g:transit_time_label>"					'[���_�ð�_��] �Ǹ��� ���� ���� ������ �ٸ� ��� �ð��� �Ҵ��ϴ� �� ������ �ǵ��� ��ǰ�� �Ҵ��ϴ� ��.
'		bufstr = bufstr & "		<g:max_handling_time>3</g:max_handling_time>"								'[�ִ�_��ǰ_�غ�_�Ⱓ] ��ǰ�� �ֹ��� �� ��۵Ǳ���� �ɸ��� �ִ� �ð��Դϴ�.
'		bufstr = bufstr & "		<g:min_handling_time>3</g:min_handling_time>"								'[�ּ�_��ǰ_�غ�_�Ⱓ] ��ǰ�� �ֹ��� �� ��۵Ǳ���� �ɸ��� �ִ� �ð��Դϴ�.
		'** ���� **
'		bufstr = bufstr & "		<g:tax>"																'#[����] �̱����ش� | �ۼ�Ʈ ������ ��ǰ �Ǹż���
'		bufstr = bufstr & "			<g:rate>5.00</g:rate>"													'#[����] �ۼ�Ʈ ������ ����
'		bufstr = bufstr & "			<g:country>US</g:country>"												'[����] ISO 3166 ���� �ڵ�
'		bufstr = bufstr & "			<g:region>MA</g:region>"												'[����]
'		bufstr = bufstr & "			<g:tax_ship>��</g:tax_ship>"											'[��ۼ�_����] ��ۺ� ������ �ΰ����� ���θ� �����մϴ�. ���Ǵ� ���� yes[��] �Ǵ� no[�ƴϿ�]�Դϴ�.
'		bufstr = bufstr & "		</g:tax>"
'		bufstr = bufstr & "		<g:tax_category>apparel</g:tax_category>"									'[����_ī�װ�] Ư�� ���� ��Ģ���� ��ǰ�� �з��ϴ� ī�װ�
		bufstr = bufstr & "</item>"
		tFile.WriteLine bufstr
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''�ۼ��ð� üũ
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
sqlStr = sqlStr & " VALUES ('googleFeed_ST')"
dbCTget.execute sqlStr

''������ ī��Ʈ
sqlStr ="[db_outmall].[dbo].[usp_Ten_Google_FeedDataCount]"
dbCTget.CommandTimeout = 120
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close

'response.write FTotCnt&"<br>"

Dim i, ArrRows, bufstr1
Dim iLastItemid : iLastItemid=9999999

If FTotCnt > 0 Then
    FTotPage = CLNG(FTotCnt / PageSize)
    If FTotPage <> (FTotCnt / PageSize) Then FTotPage = FTotPage + 1
    If (FTotPage > MaxPage) Then FTotPage = MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(appPath & FileName )

			bufstr1 = ""
			bufstr1 = bufstr1 & "<?xml version=""1.0"" encoding=""UTF-8""?>"
			bufstr1 = bufstr1 & "<rss xmlns:g=""http://base.google.com/ns/1.0"" version=""2.0"">"
			bufstr1 = bufstr1 & "	<channel>"
			bufstr1 = bufstr1 & "		<title>10x10</title>"
			bufstr1 = bufstr1 & "		<link>http://www.10x10.co.kr</link>"
			bufstr1 = bufstr1 & "		<description>10x10 Google Feed</description>"
			tFile.WriteLine bufstr1

			For i = 0 to FTotPage - 1
				ArrRows = ""
				sqlStr = "[db_outmall].[dbo].[usp_Ten_Google_FeedData] ("&i+1&", "&PageSize&", "&iLastItemid&")"
				dbCTget.CommandTimeout = 120
				rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
				If Not (rsCTget.EOF OR rsCTget.BOF) Then
					ArrRows = rsCTget.getRows()
				End If
				rsCTget.close

				If isArray(ArrRows) Then
					CALL WriteMakeGooleFeedFile(tFile, ArrRows, iLastItemid)
				End If

				''�ۼ��ð� üũ
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
				sqlStr = sqlStr & " VALUES ('googleFeed_"&(i+1)*PageSize&"_"&iLastItemid&"')"
				dbCTget.execute sqlStr
			Next
			bufstr1 = ""
			bufstr1 = bufstr1 & "	</channel>"
			bufstr1 = bufstr1 & "</rss>"
			tFile.WriteLine bufstr1
    		tFile.Close
		Set tFile = Nothing
	Set fso = Nothing
End If

''�ۼ��ð� üũ
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_nate_scraplog (ref) "
sqlStr = sqlStr & " VALUES ('googleFeed_ED')"
dbCTget.execute sqlStr

Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"�� ���� ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->