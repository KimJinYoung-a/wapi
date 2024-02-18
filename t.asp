<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 미출고상품 품절상품/출고지연 안내
' History : 이상구 생성
'           2020.10.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
dim kakaomsgstr, btnJson, smstitlestr, smsmsgstr, orderserial, itemName, buyhp
orderserial="01234567891"
itemName="test상품명"
buyhp="010-9177-8708"
						    ' smstitlestr = "[텐바이텐]주문하신 상품 택배파업 배송불가 안내"
							' smsmsgstr = "[10x10] 택배파업 배송불가안내" & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "죄송합니다. 고객님" & vbCrLf
							' smsmsgstr = smsmsgstr & "택배파업으로 인해 고객님의 배송지로 택배발송이 어렵게되어 안내드립니다." & vbCrLf
							' smsmsgstr = smsmsgstr & "현재 배송재개 가능 일정을 알수없는 상황으로" & vbCrLf
							' smsmsgstr = smsmsgstr & "안타깝지만 주문취소 안내드리는 점 양해부탁드립니다." & vbCrLf
							' smsmsgstr = smsmsgstr & "주문상품은 익일 자동 취소 및 환불예정입니다." & vbCrLf & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							' smsmsgstr = smsmsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							' smsmsgstr = smsmsgstr & "감사합니다." & vbCrLf
							' smsmsgstr = smsmsgstr & "취소하기 : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							' btnJson = "{""button"":[{""name"":""취소하기"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							' kakaomsgstr = "[10x10] 택배파업 배송불가안내" & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "죄송합니다. 고객님" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "택배파업으로 인해 고객님의 배송지로 택배발송이 어렵게되어 안내드립니다." & vbCrLf
							' kakaomsgstr = kakaomsgstr & "현재 배송재개 가능 일정을 알수없는 상황으로" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "안타깝지만 주문취소 안내드리는 점 양해부탁드립니다." & vbCrLf
							' kakaomsgstr = kakaomsgstr & "주문상품은 익일 자동 취소 및 환불예정입니다." & vbCrLf & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							' kakaomsgstr = kakaomsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							' kakaomsgstr = kakaomsgstr & "감사합니다."
							' Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0025",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")
' call SendNormalLMS("01091778708", "제목1", "1644-6030", "내용1")
' call SendNormalLMSTimeFix("01091778708", "제목2", "1644-6030", "내용2")
' call SendNormalSMS_LINK("01091778708", "1644-6030", "내용3")
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
