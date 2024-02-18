<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 카카오기프트 Regitem에 등록..젠킨스 배치처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteCSOrderLib.asp"-->
<%
Dim sqlStr
sqlStr = ""
sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_kakaoGift_regItem (itemid ,reguserid,addDlvPrice,addKakaoPrice ,kakaoitemname,kakaoGiftRegdate "
sqlStr = sqlStr & " ,kakaoGiftLastUpdate,kakaoGiftGoodNo,kakaoGiftGoodNo2,kakaoGiftPrice,kakaosaleprice,kakaoGiftSellYn, "
sqlStr = sqlStr & " kakaoGiftStatCd,regitemname,kakaostate)  "
sqlStr = sqlStr & " SELECT SellNumber,'xapi',0,0,ProductName,getdate(),getdate(),kakaono,'',NormalPrice,0,'Y',7,ProductName,'2' "
sqlStr = sqlStr & " FROM db_temp.dbo.Tbl_kakaotemp  "
sqlStr = sqlStr & " WHERE SellNumber  in ( "
sqlStr = sqlStr & " 	SELECT SellNumber  "
sqlStr = sqlStr & " 	FROM db_temp.dbo.Tbl_kakaotemp  "
sqlStr = sqlStr & " 	WHERE convert(varchar(20),SellNumber) not in  "
sqlStr = sqlStr & " 	( "
sqlStr = sqlStr & " 		SELECT convert(varchar(20),itemid) FROM db_etcmall.dbo.tbl_kakaoGift_regItem  "
sqlStr = sqlStr & " 	) "
sqlStr = sqlStr & " ) "
sqlStr = sqlStr & " and SellNumber not like '%c%' "
dbget.Execute sqlStr

rw "OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
