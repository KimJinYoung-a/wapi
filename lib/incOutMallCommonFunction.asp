<%

function DrawApiMallSelect(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >����"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >�Ե�����"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >�Ե�iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >������ũ"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='gseshop' "& chkIIF(selsitename="gseshop","selected","") &" >gseshop"
	buf = buf&"<option value='homeplus' "& chkIIF(selsitename="homeplus","selected","") &" >homeplus"
	buf = buf&"<option value='ezwel' "& chkIIF(selsitename="ezwel","selected","") &" >���������"
	buf = buf&"</select>"

	response.write buf
end function

function DrawApiMallSelectSongjangInput(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >����"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >�Ե�����"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >�Ե�iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >������ũ"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='shoplinker' "& chkIIF(selsitename="shoplinker","selected","") &" >shoplinker"
	buf = buf&"</select>"

	response.write buf
end function

''��𿡻��?
function DrawApiMallCheck()
    dim buf
    buf = ""
    buf = buf&"<input type='checkbox' name='outmallck' value='interpark'>������ũ"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteCom'>�Ե�����"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteimall'>�Ե�iMall"

    response.write buf
end function

function TenDlvCode2HomeplusDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2HomeplusDlvCode = "27220"     ''����
        CASE "2" : TenDlvCode2HomeplusDlvCode = "27230"     ''����
        CASE "3" : TenDlvCode2HomeplusDlvCode = "27030"     ''�������
        CASE "4" : TenDlvCode2HomeplusDlvCode = "27250"     ''CJ GLS
        CASE "5" : TenDlvCode2HomeplusDlvCode = "27150"     ''��Ŭ����
        CASE "6" : TenDlvCode2HomeplusDlvCode = "27260"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2HomeplusDlvCode = "27240"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2HomeplusDlvCode = "27120"     ''��ü���ù�
        CASE "9" : TenDlvCode2HomeplusDlvCode = "27270"     ''KGB�ù�
        CASE "10" : TenDlvCode2HomeplusDlvCode = "27090"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2HomeplusDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2HomeplusDlvCode = "27200"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2HomeplusDlvCode = "27100"     ''���ο�ĸ
        CASE "14" : TenDlvCode2HomeplusDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2HomeplusDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2HomeplusDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2HomeplusDlvCode = "27180"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2HomeplusDlvCode = "27060"     ''�����ù�
        CASE "19" : TenDlvCode2HomeplusDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2HomeplusDlvCode = "27280"     ''KT������
        CASE "21" : TenDlvCode2HomeplusDlvCode = "27010"     ''�浿�ù�
        CASE "22" : TenDlvCode2HomeplusDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2HomeplusDlvCode = "27080"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2HomeplusDlvCode = "27070"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2HomeplusDlvCode = "27190"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2HomeplusDlvCode = ""     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2HomeplusDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2HomeplusDlvCode = "27050"     ''�����ͽ�������
        CASE "29" : TenDlvCode2HomeplusDlvCode = ""     ''�ǿ��ù�
        CASE "30" : TenDlvCode2HomeplusDlvCode = "27140"     ''�̳�����
        CASE "31" : TenDlvCode2HomeplusDlvCode = "27170"     ''õ���ù�
        CASE "33" : TenDlvCode2HomeplusDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2HomeplusDlvCode = "27300"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2HomeplusDlvCode = "27290"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "98" : TenDlvCode2HomeplusDlvCode = "27160"     ''������->�����
        CASE "99" : TenDlvCode2HomeplusDlvCode = "27290"     ''��Ÿ
        CASE  Else
            TenDlvCode2HomeplusDlvCode = "27290"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2cjMallDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2cjMallDlvCode = "15"     ''����
        CASE "2" : TenDlvCode2cjMallDlvCode = "11"     ''����
        CASE "3" : TenDlvCode2cjMallDlvCode = "12"     ''�������
        CASE "4" : TenDlvCode2cjMallDlvCode = "22"     ''CJ GLS
        CASE "5" : TenDlvCode2cjMallDlvCode = "21"     ''��Ŭ����
        CASE "6" : TenDlvCode2cjMallDlvCode = "29"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2cjMallDlvCode = "79"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2cjMallDlvCode = "16"     ''��ü���ù�
        CASE "9" : TenDlvCode2cjMallDlvCode = "93"     ''KGB�ù�
        CASE "10" : TenDlvCode2cjMallDlvCode = "67"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2cjMallDlvCode = "17"     ''�������ù�
        CASE "12" : TenDlvCode2cjMallDlvCode = "99"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2cjMallDlvCode = "69"     ''���ο�ĸ
        CASE "14" : TenDlvCode2cjMallDlvCode = "99"     ''���̽��ù�
        CASE "15" : TenDlvCode2cjMallDlvCode = "99"     ''�߾��ù�
        CASE "16" : TenDlvCode2cjMallDlvCode = "99"     ''�����ù�
        CASE "17" : TenDlvCode2cjMallDlvCode = "57"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2cjMallDlvCode = "70"     ''�����ù�
        CASE "19" : TenDlvCode2cjMallDlvCode = "99"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2cjMallDlvCode = "68"     ''KT������
        CASE "21" : TenDlvCode2cjMallDlvCode = "78"     ''�浿�ù�
        CASE "22" : TenDlvCode2cjMallDlvCode = "99"     ''�����ù�
        CASE "23" : TenDlvCode2cjMallDlvCode = "99"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2cjMallDlvCode = "62"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2cjMallDlvCode = "60"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2cjMallDlvCode = "71"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2cjMallDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2cjMallDlvCode = "87"     ''�����ͽ�������
        CASE "29" : TenDlvCode2cjMallDlvCode = "65"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2cjMallDlvCode = "88"     ''�̳�����
        CASE "31" : TenDlvCode2cjMallDlvCode = "82"     ''õ���ù�
        CASE "33" : TenDlvCode2cjMallDlvCode = "58"     ''ȣ���ù�
        CASE "34" : TenDlvCode2cjMallDlvCode = "81"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2cjMallDlvCode = "99"     ''CVSnet�ù�  - 99 ��Ÿ�߼����ù�
        CASE "98" : TenDlvCode2cjMallDlvCode = "32"     ''������->�����
        CASE "99" : TenDlvCode2cjMallDlvCode = "99"     ''��Ÿ
        CASE  Else
            TenDlvCode2cjMallDlvCode = "99"      ''��Ÿ�߼�
    end Select
end function

function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2InterParkDlvCode = "169178"     ''����
        CASE "2" : TenDlvCode2InterParkDlvCode = "169198"     ''����
        CASE "3" : TenDlvCode2InterParkDlvCode = "169177"     ''�������
        CASE "4" : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE "5" : TenDlvCode2InterParkDlvCode = "169211"     ''��Ŭ����
        CASE "6" : TenDlvCode2InterParkDlvCode = "169181"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2InterParkDlvCode = "231145"     ''����(���ѹ̸�)
        CASE "8" : TenDlvCode2InterParkDlvCode = "169199"     ''��ü���ù�
        CASE "9" : TenDlvCode2InterParkDlvCode = "169187"     ''KGB�ù�
        CASE "10" : TenDlvCode2InterParkDlvCode = "169194"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2InterParkDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2InterParkDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE "13" : TenDlvCode2InterParkDlvCode = "169200"     ''���ο�ĸ
        CASE "14" : TenDlvCode2InterParkDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2InterParkDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2InterParkDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2InterParkDlvCode = ""     ''Ʈ����ù�
        CASE "18" : TenDlvCode2InterParkDlvCode = "169182"     ''�����ù�
        CASE "19" : TenDlvCode2InterParkDlvCode = ""     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2InterParkDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''�浿�ù�
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''�����ù�
        CASE "23" : TenDlvCode2InterParkDlvCode = "236288"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2InterParkDlvCode = "231491"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2InterParkDlvCode = "229381"     ''�ϳ����ù�
        CASE "26" : TenDlvCode2InterParkDlvCode = "263792"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX�ù�
        CASE "28" : TenDlvCode2InterParkDlvCode = "231145"     ''�����ͽ�������
        CASE "29" : TenDlvCode2InterParkDlvCode = "231194"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2InterParkDlvCode = "266237"     ''�̳�����
        CASE "31" : TenDlvCode2InterParkDlvCode = "230175"     ''õ���ù�
        CASE "33" : TenDlvCode2InterParkDlvCode = "250701"     ''ȣ���ù�
        CASE "34" : TenDlvCode2InterParkDlvCode = "258064"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2InterParkDlvCode = "169172"     ''CVSnet�ù�
        CASE "98" : TenDlvCode2InterParkDlvCode = "169316"     ''������->�����
        CASE "99" : TenDlvCode2InterParkDlvCode = "169167"     ''��Ÿ
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""      ''��Ÿ�߼�(169167)
    end Select
end function

function TenDlvCode2LotteDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="99"

    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2LotteDlvCode = "27"     ''����
        CASE "2" : TenDlvCode2LotteDlvCode = "1"     ''����v
        CASE "3" : TenDlvCode2LotteDlvCode = "31"     ''�������
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteDlvCode = "70"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''�����ù�
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT������
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteDlvCode = "70"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''�̳�����
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet�ù�
        CASE "39" : TenDlvCode2LotteDlvCode = "70"     ''KG������
        CASE "98" : TenDlvCode2LotteDlvCode = "99"     ''������
        CASE "99" : TenDlvCode2LotteDlvCode = "99"     ''��ü����
        CASE  Else
            TenDlvCode2LotteDlvCode = "99"
    end Select
end function


'''�Ե�iMall ���庯ȯ
function TenDlvCode2LotteiMallDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	�����ù�
''99	��Ÿ

    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallDlvCode = "15"     ''����
        CASE "2" : TenDlvCode2LotteiMallDlvCode = "11"     ''����v
        CASE "3" : TenDlvCode2LotteiMallDlvCode = "12"     ''�������
        CASE "4" : TenDlvCode2LotteiMallDlvCode = "16"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallDlvCode = ""     ''��Ŭ����
        CASE "6" : TenDlvCode2LotteiMallDlvCode = "22"     ''�Ｚ HTH
        CASE "7" : TenDlvCode2LotteiMallDlvCode = "26"     ''����(���ѹ̸�) ''Ȯ
        CASE "8" : TenDlvCode2LotteiMallDlvCode = "31"     ''��ü���ù�
        CASE "9" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB�ù�
        CASE "10" : TenDlvCode2LotteiMallDlvCode = "34"     ''�����ù� / �ο���(�� ����)
        CASE "11" : TenDlvCode2LotteiMallDlvCode = ""     ''�������ù�
        CASE "12" : TenDlvCode2LotteiMallDlvCode = "37"     ''�ѱ��ù� / �ѱ�Ư��
        CASE "13" : TenDlvCode2LotteiMallDlvCode = "32"     ''���ο�ĸ
        CASE "14" : TenDlvCode2LotteiMallDlvCode = ""     ''���̽��ù�
        CASE "15" : TenDlvCode2LotteiMallDlvCode = ""     ''�߾��ù�
        CASE "16" : TenDlvCode2LotteiMallDlvCode = ""     ''�����ù�
        CASE "17" : TenDlvCode2LotteiMallDlvCode = "36"     ''Ʈ����ù�
        CASE "18" : TenDlvCode2LotteiMallDlvCode = "24"     ''�����ù�
        CASE "19" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGBƯ���ù�
        CASE "20" : TenDlvCode2LotteiMallDlvCode = ""     ''KT������
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''�浿�ù�
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''�����ù�
        CASE "23" : TenDlvCode2LotteiMallDlvCode = "47"     ''�굦���ù� �ż���
        CASE "24" : TenDlvCode2LotteiMallDlvCode = "43"     ''�簡���ͽ�������
        CASE "25" : TenDlvCode2LotteiMallDlvCode = "46"     ''�ϳ����ù�v
        CASE "26" : TenDlvCode2LotteiMallDlvCode = "18"     ''�Ͼ��ù�
        CASE "27" : TenDlvCode2LotteiMallDlvCode = "48"     ''LOEX�ù�
        CASE "28" : TenDlvCode2LotteiMallDlvCode = "26"     ''�����ͽ�������
        CASE "29" : TenDlvCode2LotteiMallDlvCode = "99"     ''�ǿ��ù�
        CASE "30" : TenDlvCode2LotteiMallDlvCode = "23"     ''�̳�����
        CASE "31" : TenDlvCode2LotteiMallDlvCode = "17"     ''õ���ù�
        CASE "33" : TenDlvCode2LotteiMallDlvCode = ""     ''ȣ���ù�
        CASE "34" : TenDlvCode2LotteiMallDlvCode = "38"     ''���ȭ���ù�
        CASE "35" : TenDlvCode2LotteiMallDlvCode = "99"     ''CVSnet�ù�
        CASE "98" : TenDlvCode2LotteiMallDlvCode = "99"     ''������
        CASE "99" : TenDlvCode2LotteiMallDlvCode = "99"     ''��ü����
        CASE  Else
            TenDlvCode2LotteiMallDlvCode = "99"
    end Select
end function

function LotteiMallDlvCode2Name(iltDlvCode)
    LotteiMallDlvCode2Name = "��Ÿ"
    if IsNULL(iltDlvCode) then Exit function
    iltDlvCode = TRIM(CStr(iltDlvCode))

    select Case iltDlvCode
        CASE "11" : LotteiMallDlvCode2Name="�����ù�"
        CASE "12" : LotteiMallDlvCode2Name="�����̴������"
        CASE "15" : LotteiMallDlvCode2Name="�����ù�"
        CASE "16" : LotteiMallDlvCode2Name="CJGLS"
        CASE "17" : LotteiMallDlvCode2Name="õ���ù�"
        CASE "18" : LotteiMallDlvCode2Name="�Ͼ��ù�"
        CASE "19" : LotteiMallDlvCode2Name="��Ÿ�ù�"
        CASE "22" : LotteiMallDlvCode2Name="HTH�ù�"
        CASE "24" : LotteiMallDlvCode2Name="�����ù�"
        CASE "26" : LotteiMallDlvCode2Name="�����ͽ�������"
        CASE "31" : LotteiMallDlvCode2Name="��ü���ù�"
        CASE "32" : LotteiMallDlvCode2Name="���ο�ĸ"
        CASE "34" : LotteiMallDlvCode2Name="�����ù�"
        CASE "36" : LotteiMallDlvCode2Name="Ʈ���"
        CASE "37" : LotteiMallDlvCode2Name="�ѱ��ù�"
        CASE "38" : LotteiMallDlvCode2Name="����ù�"
        CASE "40" : LotteiMallDlvCode2Name="KGB�ù�"
        CASE "41" : LotteiMallDlvCode2Name="�����ù�"
        CASE "43" : LotteiMallDlvCode2Name="�簡���ͽ�������"
        CASE "46" : LotteiMallDlvCode2Name="�ϳ����ù�"
        CASE "47" : LotteiMallDlvCode2Name="�������ù�"
        CASE "48" : LotteiMallDlvCode2Name="�ο����ù�"
        CASE "49" : LotteiMallDlvCode2Name="�浿�ù�"
        CASE "99" : LotteiMallDlvCode2Name="��Ÿ"
        CASE  Else
            LotteiMallDlvCode2Name = "��Ÿ"
    end Select
end function

function Fn_ActOutMall_CateSummary(iMallID)
    dim sqlStr
    sqlStr = "exec db_item.dbo.sp_Ten_OutMall_CateSummary '"&iMallID&"'"
    dbget.Execute sqlStr

	If iMallID = "cjmall" Then
    	sqlStr = "exec db_outmall.dbo.sp_Ten_OutMall_CateSummary '"&iMallID&"'"
    	dbCTget.Execute sqlStr
    End If
end function

Function Fn_AcctFailTouch(iMallID,iitemid,iLastErrStr)
    Dim strSql
    iLastErrStr = html2db(iLastErrStr)

    IF (iMallID="lotteCom") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_lotte_regItem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)

    ELSEIF (iMallID="lotteimall") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_LTiMall_regItem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)

    ELSEIF (iMallID="interpark") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_interpark_reg_item R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)
    ELSEIF (iMallID="gsshop") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_gsshop_regitem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)
    ELSEIF (iMallID="homeplus") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_outmall.dbo.tbl_homeplus_regitem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbCTget.Execute(strSql)
    ELSEIF (iMallID="ezwel") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_outmall.dbo.tbl_ezwel_regitem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbCTget.Execute(strSql)
    ENd IF

end function


function Fn_AcctFailLog(iMallID,iitemid,ErrMsg,ErrCode)
    Dim sqlStr
    ''db_log.dbo.tbl_interparkEdit_log
    IF (iMallID="lotteCom") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(lotteGoodNo,lotteTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_lotte_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ELSEIF (iMallID="lotteimall") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.LTimallGoodno,R.LtiMallTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_ltimall_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr

    ELSEIF (iMallID="interpark") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.interparkprdno,''), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ENd IF
end function

function Fn_AcctFailLogNone(iMallID,iitemid,ioutmallPrdno,ioutmallsellyn,ioutmallsellcash,ioutmallbuycash,ErrMsg,ErrCode)
    Dim sqlStr
    sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
    sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
    sqlStr = sqlStr & " values("&iitemid& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallPrdno&"'"& VbCrlf
    sqlStr = sqlStr & " ,"&ioutmallsellcash& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallbuycash&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallsellyn&"'"& VbCrlf
    sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(ErrMsg)&"')"& VbCrlf
    sqlStr = sqlStr & " ,'"&ErrCode&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&iMallID&"')"& VbCrlf
    dbget.execute sqlStr
end function
%>