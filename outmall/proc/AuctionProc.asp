<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/outmall/auction/incAuctionFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, arrRows, skipItem, oAuction, getMustprice, tAuctionGoodno, oAuctionOpt
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, isiframe, isAllRegYn
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
If itemid="" or itemid="0" Then
	response.write "<script>alert('상품번호가 없습니다.')</script>"
	response.end
ElseIf Not(isNumeric(itemid)) Then
	response.write "<script>alert('잘못된 상품번호입니다.')</script>"
	response.end
Else
	'정수형태로 변환
	itemid=CLng(getNumeric(itemid))
End If
'######################################################## Auction API ########################################################
If mallid = "auction1010" Then
	If action = "REG" Then					'상품등록
		'##################################### 기본 정보 등록 시작 #####################################
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionNotRegOneItem
		    If (oAuction.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf (oAuction.FOneItem.FNotinCate = "Y") Then
				iErrStr = "ERR||"&itemid&"||상품 등록 제외 카테고리입니다."
			Else
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
				strSql = strSql & " BEGIN"& VbCRLF
				strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem "
		        strSql = strSql & " 	(itemid, regdate, reguserid, auctionstatCD, regitemname, auctionSellYn)"
		        strSql = strSql & " 	VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oAuction.FOneItem.FItemName)&"', 'N')"
				strSql = strSql & " END "
				dbget.Execute strSql

				'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
				If oAuction.FOneItem.checkTenItemOptionValid Then
					strParam = ""
					strParam = oAuction.FOneItem.getAuctionItemRegParameter()
					getMustprice = ""
					getMustprice = oAuction.FOneItem.MustPrice()
					Call fnAuctionItemReg(itemid, strParam, iErrStr, getMustprice, oAuction.FOneItem.getAuctionSellYn, oAuction.FOneItem.FLimityn, oAuction.FOneItem.FLimitNo, oAuction.FOneItem.FLimitSold, html2db(oAuction.FOneItem.FItemName), oAuction.FOneItem.FbasicimageNm)
				Else
					iErrStr = "ERR||"&itemid&"||[AddItem] 옵션검사 실패"
				End If
			End If
		SET oAuction = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
		'##################################### 기본 정보 등록 끝 #####################################

		'#################################### 옵션 정보 등록 시작 ####################################
		If failCnt = 0 Then
			SET oAuctionOpt = new CAuction
				oAuctionOpt.FRectItemID	= itemid
				oAuctionOpt.getAuctionNotOptOneItem
			    If (oAuctionOpt.FResultCount < 1) Then
					iErrStr = "ERR||"&itemid&"||옵션 등록 가능한 상품이 아닙니다."
				ElseIf (oAuctionOpt.FOneItem.FAuctionGoodNo = "") Then
					iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
				ElseIf (oAuctionOpt.FOneItem.FAPIadditem = "N") Then
					iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
				ElseIf (oAuctionOpt.FOneItem.FAPIaddopt = "Y") Then
					iErrStr = "ERR||"&itemid&"||이미 옵션정보를 등록하셨습니다."
				Else
					strParam = ""
					strParam = oAuctionOpt.FOneItem.getAuctionOPTRegParameter()
					Call fnAuctionOPTReg(itemid, strParam, iErrStr)
				End If
				tAuctionGoodno = oAuctionOpt.FOneItem.FAuctionGoodNo
			SET oAuctionOpt = nothing
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'#################################### 옵션 정보 등록 끝 ####################################

		'################################# 상품고시 정보 등록 시작 #################################
		If failCnt = 0 Then
			If tAuctionGoodno = "" Then
				iErrStr = "ERR||"&itemid&"||기본정보부터 입력하셔야 됩니다."
			Else
				strParam = ""
				strParam = getAuctionInfoCdParameter(itemid, tAuctionGoodno)
				Call fnAuctionItemInfoCd(itemid, strParam, iErrStr)
			End If
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
		'################################## 상품고시 정보 등록 끝 ##################################
		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=REG
	ElseIf action = "REGOnSale" Then						'옵션 조회 후  신규등록 상품 판매중으로 변경
		isAllRegYn = getAllRegChk(itemid)
		If isAllRegYn <> "Y" Then
			iErrStr = "ERR||"&itemid&"||기본정보, 옵션정보, 상품고시 입력을 확인하세요"
		Else
			tAuctionGoodno = getAuctionGoodno(itemid)
			strParam = ""
			strParam = getAuctionOptSellModParameter(tAuctionGoodno)
			Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				strParam = ""
				strParam = getAuctionSellYnParameter("Y", itemid, tAuctionGoodno)
				Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=REGOnSale
	ElseIf action = "SOLDOUT" Then			'상태변경
		strParam = ""
		strParam = getAuctionSellYnParameter("N", itemid, getAuctionGoodno(itemid))
		Call fnAuctionSellyn(itemid, "N", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=SOLDOUT
	ElseIf action = "KEEPSELL" Then		'상품 판매 유지
		strParam = ""
		strParam = getAuctionSellYnParameter("Y", itemid, getAuctionGoodno(itemid))
		Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
		lastErrStr = iErrStr
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=KEEPSELL
	ElseIf action = "PRICE" Then		'가격수정
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionEditOneItem
		    If (oAuction.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			ElseIf getAllRegChk2(itemid) <> "Y" Then
				iErrStr = "ERR||"&itemid&"||OnSale변경 확인하세요"
			Else
				strParam = ""
				strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()
				getMustprice = ""
				getMustprice = oAuction.FOneItem.MustPrice()
				Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
			End If

			If (Left(iErrStr,2)) <> "OK" and (Left(iErrStr,2)) <> "ER" Then
				iErrStr = "ERR||"&itemid&"||잘못된 호출"
			End If
			lastErrStr = iErrStr
			response.write iErrStr
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("auction1010", itemid, iErrStr)
			End If
		SET oAuction = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=PRICE
	ElseIf action = "EDIT" Then			'재고조회 + 상품정보 + 가격 + 필요에 따라 상품판매상태수정
		SET oAuction = new CAuction
			oAuction.FRectItemID	= itemid
			oAuction.getAuctionEditOneItem
			If oAuction.FResultCount > 0 Then
				If oAuction.FOneItem.checkItemContent = "Y" Then
					isiframe = "Y"
				End If

				If (oAuction.FOneItem.FmaySoldOut = "Y") OR (isiframe = "Y") OR (oAuction.FOneItem.IsMayLimitSoldout = "Y") Then
					strParam = ""
					strParam = getAuctionSellYnParameter("N", itemid, oAuction.FOneItem.FAuctionGoodNo)
					Call fnAuctionSellyn(itemid, "N", strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				Else
					If (oAuction.FOneItem.FAuctionSellYn = "N" AND oAuction.FOneItem.IsSoldOut = False) Then
						iErrStr = ""
						strParam = ""
						strParam = getAuctionSellYnParameter("Y", itemid, oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionSellyn(itemid, "Y", strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If

					strParam = ""
					strParam = oAuction.FOneItem.getAuctionItemInfoEditParameter()
					getMustprice = ""
					getMustprice = oAuction.FOneItem.MustPrice()
					Call fnAuctionIteminfoEdit(itemid, oAuction.FOneItem.FAuctionGoodNo, iErrStr, strParam, getMustprice)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					If oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then			'텐바이텐 옵션있고, 옥션의 옵션도 등록되어있다면..즉 둘다 옵션상태
						'## 총 3번의 API를 돌려야 될 것
						'1.옵션을 완전 초기화
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
						Call fnAuctionOPTDel(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'2.초기화 후 재 세팅
						strParam = ""
						strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
						Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If

						'3.옵션 조회로 가져오기
						strParam = ""
						strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
						Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					Else
						If oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then		'둘 다 단품 상태
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionDanPoomModParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. 옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						ElseIf oAuction.FOneItem.FOptioncnt > 0 AND oAuction.FOneItem.FRegedoptcnt = 0 Then	'텐바이텐상품이 옵션있음 으로 변경되고, 등록된 옵션은 없는 상태
							'1. 재 세팅
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTRegParameter()
							Call fnAuctionOPTEDT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							If failCnt = 0 Then
								strSql = ""
								strSql = " DELETE FROM db_item.dbo.tbl_outmall_regedoption WHERE mallid = '"&CMALLNAME&"' and itemid = '"&itemid&"' "
								dbget.Execute(strSql)

								strSql = ""
								strSql = "UPDATE db_etcmall.dbo.tbl_auction_regitem SET regedoptcnt = null WHERE itemid = '"&itemid&"'"
								dbget.Execute(strSql)
							End If

							'2. 옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

						ElseIf oAuction.FOneItem.FOptioncnt = 0 AND oAuction.FOneItem.FRegedoptcnt > 0 Then	'텐바이텐상품이 옵션있음에서 단품으로 변경되고, 등록된 옵션이 있는 상태
							'1.옵션을 완전 초기화
							strParam = ""
							strParam = oAuction.FOneItem.getAuctionOPTDeleteParameter()
							Call fnAuctionOPTDel(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If

							'2.옵션 조회로 가져오기
							strParam = ""
							strParam = getAuctionOptSellModParameter(oAuction.FOneItem.FAuctionGoodNo)
							Call fnAuctionOPTSTAT(itemid, strParam, iErrStr)
							If Left(iErrStr, 2) <> "OK" Then
								failCnt = failCnt + 1
								SumErrStr = SumErrStr & iErrStr
							Else
								SumOKStr = SumOKStr & iErrStr
							End If
						End If
					End If
				End If

				'OK던 ERR이던 editQuecnt에 + 1을 시킴..
				'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
				strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
				strSql = strSql & " ,AuctionLastUpdate = getdate()  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("auction1010", itemid, SumErrStr)
					lastErrStr = "ERR||"&itemid&"||"&SumErrStr
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_auction_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql

					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					lastErrStr = "OK||"&itemid&"||"&SumOKStr
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET oAuction = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/AuctionProc.asp?itemid=1159694&mallid=auction1010&action=EDIT
	End If
End If
'###################################################### Auction API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->