<%

function DrawApiMallSelect(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >선택"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >롯데닷컴"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >롯데iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >인터파크"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='gseshop' "& chkIIF(selsitename="gseshop","selected","") &" >gseshop"
	buf = buf&"<option value='homeplus' "& chkIIF(selsitename="homeplus","selected","") &" >homeplus"
	buf = buf&"<option value='ezwel' "& chkIIF(selsitename="ezwel","selected","") &" >이지웰페어"
	buf = buf&"</select>"

	response.write buf
end function

function DrawApiMallSelectSongjangInput(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >선택"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >롯데닷컴"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >롯데iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >인터파크"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='shoplinker' "& chkIIF(selsitename="shoplinker","selected","") &" >shoplinker"
	buf = buf&"</select>"

	response.write buf
end function

''어디에사용?
function DrawApiMallCheck()
    dim buf
    buf = ""
    buf = buf&"<input type='checkbox' name='outmallck' value='interpark'>인터파크"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteCom'>롯데닷컴"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteimall'>롯데iMall"

    response.write buf
end function

function TenDlvCode2HomeplusDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2HomeplusDlvCode = "27220"     ''한진
        CASE "2" : TenDlvCode2HomeplusDlvCode = "27230"     ''현대
        CASE "3" : TenDlvCode2HomeplusDlvCode = "27030"     ''대한통운
        CASE "4" : TenDlvCode2HomeplusDlvCode = "27250"     ''CJ GLS
        CASE "5" : TenDlvCode2HomeplusDlvCode = "27150"     ''이클라인
        CASE "6" : TenDlvCode2HomeplusDlvCode = "27260"     ''삼성 HTH
        CASE "7" : TenDlvCode2HomeplusDlvCode = "27240"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2HomeplusDlvCode = "27120"     ''우체국택배
        CASE "9" : TenDlvCode2HomeplusDlvCode = "27270"     ''KGB택배
        CASE "10" : TenDlvCode2HomeplusDlvCode = "27090"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2HomeplusDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2HomeplusDlvCode = "27200"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2HomeplusDlvCode = "27100"     ''옐로우캡
        CASE "14" : TenDlvCode2HomeplusDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2HomeplusDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2HomeplusDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2HomeplusDlvCode = "27180"     ''트라넷택배
        CASE "18" : TenDlvCode2HomeplusDlvCode = "27060"     ''로젠택배
        CASE "19" : TenDlvCode2HomeplusDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2HomeplusDlvCode = "27280"     ''KT로지스
        CASE "21" : TenDlvCode2HomeplusDlvCode = "27010"     ''경동택배
        CASE "22" : TenDlvCode2HomeplusDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2HomeplusDlvCode = "27080"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2HomeplusDlvCode = "27070"     ''사가와익스프레스
        CASE "25" : TenDlvCode2HomeplusDlvCode = "27190"     ''하나로택배
        CASE "26" : TenDlvCode2HomeplusDlvCode = ""     ''일양택배
        CASE "27" : TenDlvCode2HomeplusDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2HomeplusDlvCode = "27050"     ''동부익스프레스
        CASE "29" : TenDlvCode2HomeplusDlvCode = ""     ''건영택배
        CASE "30" : TenDlvCode2HomeplusDlvCode = "27140"     ''이노지스
        CASE "31" : TenDlvCode2HomeplusDlvCode = "27170"     ''천일택배
        CASE "33" : TenDlvCode2HomeplusDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2HomeplusDlvCode = "27300"     ''대신화물택배
        CASE "35" : TenDlvCode2HomeplusDlvCode = "27290"     ''CVSnet택배  - 99 기타중소형택배
        CASE "98" : TenDlvCode2HomeplusDlvCode = "27160"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2HomeplusDlvCode = "27290"     ''기타
        CASE  Else
            TenDlvCode2HomeplusDlvCode = "27290"      ''기타발송
    end Select
end function

function TenDlvCode2cjMallDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2cjMallDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2cjMallDlvCode = "11"     ''현대
        CASE "3" : TenDlvCode2cjMallDlvCode = "12"     ''대한통운
        CASE "4" : TenDlvCode2cjMallDlvCode = "22"     ''CJ GLS
        CASE "5" : TenDlvCode2cjMallDlvCode = "21"     ''이클라인
        CASE "6" : TenDlvCode2cjMallDlvCode = "29"     ''삼성 HTH
        CASE "7" : TenDlvCode2cjMallDlvCode = "79"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2cjMallDlvCode = "16"     ''우체국택배
        CASE "9" : TenDlvCode2cjMallDlvCode = "93"     ''KGB택배
        CASE "10" : TenDlvCode2cjMallDlvCode = "67"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2cjMallDlvCode = "17"     ''오렌지택배
        CASE "12" : TenDlvCode2cjMallDlvCode = "99"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2cjMallDlvCode = "69"     ''옐로우캡
        CASE "14" : TenDlvCode2cjMallDlvCode = "99"     ''나이스택배
        CASE "15" : TenDlvCode2cjMallDlvCode = "99"     ''중앙택배
        CASE "16" : TenDlvCode2cjMallDlvCode = "99"     ''주코택배
        CASE "17" : TenDlvCode2cjMallDlvCode = "57"     ''트라넷택배
        CASE "18" : TenDlvCode2cjMallDlvCode = "70"     ''로젠택배
        CASE "19" : TenDlvCode2cjMallDlvCode = "99"     ''KGB특급택배
        CASE "20" : TenDlvCode2cjMallDlvCode = "68"     ''KT로지스
        CASE "21" : TenDlvCode2cjMallDlvCode = "78"     ''경동택배
        CASE "22" : TenDlvCode2cjMallDlvCode = "99"     ''고려택배
        CASE "23" : TenDlvCode2cjMallDlvCode = "99"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2cjMallDlvCode = "62"     ''사가와익스프레스
        CASE "25" : TenDlvCode2cjMallDlvCode = "60"     ''하나로택배
        CASE "26" : TenDlvCode2cjMallDlvCode = "71"     ''일양택배
        CASE "27" : TenDlvCode2cjMallDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2cjMallDlvCode = "87"     ''동부익스프레스
        CASE "29" : TenDlvCode2cjMallDlvCode = "65"     ''건영택배
        CASE "30" : TenDlvCode2cjMallDlvCode = "88"     ''이노지스
        CASE "31" : TenDlvCode2cjMallDlvCode = "82"     ''천일택배
        CASE "33" : TenDlvCode2cjMallDlvCode = "58"     ''호남택배
        CASE "34" : TenDlvCode2cjMallDlvCode = "81"     ''대신화물택배
        CASE "35" : TenDlvCode2cjMallDlvCode = "99"     ''CVSnet택배  - 99 기타중소형택배
        CASE "98" : TenDlvCode2cjMallDlvCode = "32"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2cjMallDlvCode = "99"     ''기타
        CASE  Else
            TenDlvCode2cjMallDlvCode = "99"      ''기타발송
    end Select
end function

function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2InterParkDlvCode = "169178"     ''한진
        CASE "2" : TenDlvCode2InterParkDlvCode = "169198"     ''현대
        CASE "3" : TenDlvCode2InterParkDlvCode = "169177"     ''대한통운
        CASE "4" : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE "5" : TenDlvCode2InterParkDlvCode = "169211"     ''이클라인
        CASE "6" : TenDlvCode2InterParkDlvCode = "169181"     ''삼성 HTH
        CASE "7" : TenDlvCode2InterParkDlvCode = "231145"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2InterParkDlvCode = "169199"     ''우체국택배
        CASE "9" : TenDlvCode2InterParkDlvCode = "169187"     ''KGB택배
        CASE "10" : TenDlvCode2InterParkDlvCode = "169194"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2InterParkDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2InterParkDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2InterParkDlvCode = "169200"     ''옐로우캡
        CASE "14" : TenDlvCode2InterParkDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2InterParkDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2InterParkDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2InterParkDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2InterParkDlvCode = "169182"     ''로젠택배
        CASE "19" : TenDlvCode2InterParkDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2InterParkDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''경동택배
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''고려택배
        CASE "23" : TenDlvCode2InterParkDlvCode = "236288"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2InterParkDlvCode = "231491"     ''사가와익스프레스
        CASE "25" : TenDlvCode2InterParkDlvCode = "229381"     ''하나로택배
        CASE "26" : TenDlvCode2InterParkDlvCode = "263792"     ''일양택배
        CASE "27" : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX택배
        CASE "28" : TenDlvCode2InterParkDlvCode = "231145"     ''동부익스프레스
        CASE "29" : TenDlvCode2InterParkDlvCode = "231194"     ''건영택배
        CASE "30" : TenDlvCode2InterParkDlvCode = "266237"     ''이노지스
        CASE "31" : TenDlvCode2InterParkDlvCode = "230175"     ''천일택배
        CASE "33" : TenDlvCode2InterParkDlvCode = "250701"     ''호남택배
        CASE "34" : TenDlvCode2InterParkDlvCode = "258064"     ''대신화물택배
        CASE "35" : TenDlvCode2InterParkDlvCode = "169172"     ''CVSnet택배
        CASE "98" : TenDlvCode2InterParkDlvCode = "169316"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2InterParkDlvCode = "169167"     ''기타
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""      ''기타발송(169167)
    end Select
end function

function TenDlvCode2LotteDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="99"

    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2LotteDlvCode = "27"     ''한진
        CASE "2" : TenDlvCode2LotteDlvCode = "1"     ''현대v
        CASE "3" : TenDlvCode2LotteDlvCode = "31"     ''대한통운
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''이클라인
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''우체국택배
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB택배
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteDlvCode = "70"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''중앙택배
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''로젠택배
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT로지스
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''경동택배
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''일양택배
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2LotteDlvCode = "70"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''건영택배
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''이노지스
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''천일택배
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''호남택배
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet택배
        CASE "39" : TenDlvCode2LotteDlvCode = "70"     ''KG로지스
        CASE "98" : TenDlvCode2LotteDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteDlvCode = "99"
    end Select
end function


'''롯데iMall 송장변환
function TenDlvCode2LotteiMallDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	이젠택배
''99	기타

    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2LotteiMallDlvCode = "11"     ''현대v
        CASE "3" : TenDlvCode2LotteiMallDlvCode = "12"     ''대한통운
        CASE "4" : TenDlvCode2LotteiMallDlvCode = "16"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2LotteiMallDlvCode = "22"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteiMallDlvCode = "31"     ''우체국택배
        CASE "9" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB택배
        CASE "10" : TenDlvCode2LotteiMallDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteiMallDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteiMallDlvCode = "37"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteiMallDlvCode = "32"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteiMallDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteiMallDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2LotteiMallDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteiMallDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteiMallDlvCode = "24"     ''로젠택배
        CASE "19" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteiMallDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''경동택배
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteiMallDlvCode = "47"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteiMallDlvCode = "43"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteiMallDlvCode = "46"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteiMallDlvCode = "18"     ''일양택배
        CASE "27" : TenDlvCode2LotteiMallDlvCode = "48"     ''LOEX택배
        CASE "28" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteiMallDlvCode = "99"     ''건영택배
        CASE "30" : TenDlvCode2LotteiMallDlvCode = "23"     ''이노지스
        CASE "31" : TenDlvCode2LotteiMallDlvCode = "17"     ''천일택배
        CASE "33" : TenDlvCode2LotteiMallDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2LotteiMallDlvCode = "38"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteiMallDlvCode = "99"     ''CVSnet택배
        CASE "98" : TenDlvCode2LotteiMallDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteiMallDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteiMallDlvCode = "99"
    end Select
end function

function LotteiMallDlvCode2Name(iltDlvCode)
    LotteiMallDlvCode2Name = "기타"
    if IsNULL(iltDlvCode) then Exit function
    iltDlvCode = TRIM(CStr(iltDlvCode))

    select Case iltDlvCode
        CASE "11" : LotteiMallDlvCode2Name="현대택배"
        CASE "12" : LotteiMallDlvCode2Name="씨제이대한통운"
        CASE "15" : LotteiMallDlvCode2Name="한진택배"
        CASE "16" : LotteiMallDlvCode2Name="CJGLS"
        CASE "17" : LotteiMallDlvCode2Name="천일택배"
        CASE "18" : LotteiMallDlvCode2Name="일양택배"
        CASE "19" : LotteiMallDlvCode2Name="기타택배"
        CASE "22" : LotteiMallDlvCode2Name="HTH택배"
        CASE "24" : LotteiMallDlvCode2Name="로젠택배"
        CASE "26" : LotteiMallDlvCode2Name="동부익스프레스"
        CASE "31" : LotteiMallDlvCode2Name="우체국택배"
        CASE "32" : LotteiMallDlvCode2Name="옐로우캡"
        CASE "34" : LotteiMallDlvCode2Name="아주택배"
        CASE "36" : LotteiMallDlvCode2Name="트라넷"
        CASE "37" : LotteiMallDlvCode2Name="한국택배"
        CASE "38" : LotteiMallDlvCode2Name="대신택배"
        CASE "40" : LotteiMallDlvCode2Name="KGB택배"
        CASE "41" : LotteiMallDlvCode2Name="이젠택배"
        CASE "43" : LotteiMallDlvCode2Name="사가와익스프레스"
        CASE "46" : LotteiMallDlvCode2Name="하나로택배"
        CASE "47" : LotteiMallDlvCode2Name="세덱스택배"
        CASE "48" : LotteiMallDlvCode2Name="로엑스택배"
        CASE "49" : LotteiMallDlvCode2Name="경동택배"
        CASE "99" : LotteiMallDlvCode2Name="기타"
        CASE  Else
            LotteiMallDlvCode2Name = "기타"
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