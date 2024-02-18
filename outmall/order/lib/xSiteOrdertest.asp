<%

Class OrderItem
	public FSellSite
	public FOutMallOrderSerial
End Class

class COrderMasterItem
	public FSellSite
	public FOutMallOrderSerial
	public FSellDate
	public FPayType
	public FPaydate
	public FOrderUserID
	public FOrderName
	public FOrderEmail
	public FOrderTelNo
	public FOrderHpNo
	public FReceiveName
	public FReceiveTelNo
	public FReceiveHpNo
	public FReceiveZipCode
	public FReceiveAddr1
	public FReceiveAddr2
	public Fdeliverymemo
	public FdeliverPay

	public FUserID
	public ForderCsGbn
	public FcountryCode
	public Fshoplinkermallname
	public FshoplinkerOrderID
	public FshoplinkerMallID
	public FoverseasPrice
	public FoverseasDeliveryPrice
	public FoverseasRealPrice
	public Freserve01
	public FbeasongNum11st

	Private Sub Class_Initialize()
		ForderCsGbn = 0
		FcountryCode = "KR"
		''FoverseasPrice = 0
		''FoverseasDeliveryPrice = 0
		''FoverseasRealPrice = 0
	End Sub
end class

class COrderDetail
	public FdetailSeq
	public FItemID
	public FItemOption
	public FOutMallItemID
	public FOutMallItemName
	public FOutMallItemOption
	public FOutMallItemOptionName
	public Fitemcost
	public FReducedPrice
	public FItemNo
	public FOutMallCouponPrice
	public FTenCouponPrice
	public FrequireDetail

	public FshoplinkerPrdCode
end class

function GetOrderFromExtSite(sellsite, selldate, chgCode)
	select case sellsite
		case "nvstorefarm"
			Call GetOrderFrom_nvstorefarm(sellsite, selldate, chgCode)
		case else
			response.write "잘못된 접근입니다."
		dbget.close : response.end
	end select
end function

function SaveOrderToDB(oMaster, oDetailArr)
	dim sqlStr
	dim i, j, k
	dim paramInfo, retParamInfo, RetErr, retErrStr
	dim orderDlvPay
	dim tmpStr

	SaveOrderToDB = False

	if NOT isNULL(oMaster.FReceiveZipCode) then
		if (LEN(replace(Trim(oMaster.FReceiveZipCode),"-",""))=5) then  ''5자리 우편번호이면
			oMaster.FReceiveZipCode = replace(Trim(oMaster.FReceiveZipCode),"-","")
		end if
	end if

	for i = 0 to UBound(oDetailArr)
		if (i = 0) then
			orderDlvPay = oMaster.FdeliverPay
		else
			orderDlvPay = 0
		end if

		tmpStr = " exec db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_TEST "
		tmpStr = tmpStr + "'" & oMaster.FSellSite & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOutMallOrderSerial & "'"
		tmpStr = tmpStr + ", '" & oMaster.FSellDate & "'"
		tmpStr = tmpStr + ", '" & oMaster.FPayType & "'"
		tmpStr = tmpStr + ", '" & oMaster.FPaydate & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemID & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemOption & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemID & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemName & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemOption & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemOptionName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FUserID & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderEmail & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderTelNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderHpNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveTelNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderHpNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveZipCode & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveAddr1 & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveAddr2 & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).Fitemcost & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FReducedPrice & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemNo & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FdetailSeq & "'"
		tmpStr = tmpStr + ", '" & 0 & "'"
		tmpStr = tmpStr + ", '" & 0 & "'"
		tmpStr = tmpStr + ", '" & oMaster.Fdeliverymemo & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FrequireDetail & "'"
		tmpStr = tmpStr + ", '" & orderDlvPay & "'"
		tmpStr = tmpStr + ", '" & oMaster.ForderCsGbn & "'"
		tmpStr = tmpStr + ", '" & oMaster.FcountryCode & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemID & "'"
		tmpStr = tmpStr + ", '" & oMaster.Fshoplinkermallname & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FshoplinkerPrdCode & "'"
		tmpStr = tmpStr + ", '" & oMaster.FshoplinkerOrderID & "'"
		tmpStr = tmpStr + ", '" & oMaster.FshoplinkerMallID & "'"
		tmpStr = tmpStr + ", ''"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasDeliveryPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasRealPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.Freserve01 & "'"
		tmpStr = tmpStr + ", '" & oMaster.FbeasongNum11st & "'"

		tmpStr = Replace(tmpStr, "'", "^")

		sqlStr = "insert into db_temp.dbo.tbl_tmp_gsOrder"
		sqlStr = sqlStr&" (regdate,refip,xmlData)"
		sqlStr = sqlStr&" values(getdate(),'XXX','KKK-" & tmpStr & "')"
		''dbget.Execute sqlStr

		paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
        	,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim(oMaster.FSellSite))	_
			,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(oMaster.FOutMallOrderSerial)) _
			,Array("@SellDate"				, adDate		, adParamInput		,	  , Trim(oMaster.FSellDate)) _
			,Array("@PayType"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FPayType)) _
			,Array("@Paydate"				, adDate		, adParamInput		,     , Trim(oMaster.FPaydate)) _
			,Array("@matchItemID"			, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemID)) _
			,Array("@matchItemOption"		, adVarchar		, adParamInput		,    4, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerItemID"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FItemID)) _
			,Array("@partnerItemName"		, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FOutMallItemName)) _
			,Array("@partnerOption"			, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerOptionName"		, adVarchar		, adParamInput		, 1024, Trim(oDetailArr(i).FOutMallItemOptionName)) _
			,Array("@OrderUserID"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FUserID)) _
			,Array("@OrderName"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FOrderName)) _
			,Array("@OrderEmail"			, adVarchar		, adParamInput		,  100, Trim(oMaster.FOrderEmail)) _
			,Array("@OrderTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderTelNo)) _
			,Array("@OrderHpNo"				, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderHpNo)) _
			,Array("@ReceiveName"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FReceiveName)) _
			,Array("@ReceiveTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveTelNo)) _
			,Array("@ReceiveHpNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveHpNo)) _
			,Array("@ReceiveZipCode"		, adVarchar		, adParamInput		,   20, Trim(oMaster.FReceiveZipCode)) _
			,Array("@ReceiveAddr1"			, adVarchar		, adParamInput		,  128, Trim(oMaster.FReceiveAddr1)) _
			,Array("@ReceiveAddr2"			, adVarchar		, adParamInput		,  512, Trim(oMaster.FReceiveAddr2)) _
			,Array("@SellPrice"				, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).Fitemcost)) _
			,Array("@RealSellPrice"			, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).FReducedPrice)) _
			,Array("@ItemOrderCount"		, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemNo)) _
			,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FdetailSeq)) _
			,Array("@DeliveryType"			, adInteger		, adParamInput		,     , 0) _
			,Array("@deliveryprice"			, adCurrency	, adParamInput		,     , 0) _
			,Array("@deliverymemo"			, adVarchar		, adParamInput		,  400, Trim(oMaster.Fdeliverymemo)) _
			,Array("@requireDetail"			, adVarchar		, adParamInput		, 400, Trim(oDetailArr(i).FrequireDetail)) _
			,Array("@orderDlvPay"			, adCurrency	, adParamInput		,     , orderDlvPay) _
			,Array("@orderCsGbn"			, adInteger		, adParamInput		,     , oMaster.ForderCsGbn) _
			,Array("@countryCode"			, adVarchar		, adParamInput		,    2, oMaster.FcountryCode) _
            ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(oDetailArr(i).FOutMallItemID)) _
			,Array("@shoplinkerMallName" 	, adVarchar		, adParamInput		,   64, oMaster.Fshoplinkermallname) _
			,Array("@shoplinkerPrdCode"		, adVarchar		, adParamInput		,   16, oDetailArr(i).FshoplinkerPrdCode) _
			,Array("@shoplinkerOrderID"		, adVarchar		, adParamInput		,   16, oMaster.FshoplinkerOrderID) _
			,Array("@shoplinkerMallID"		, adVarchar		, adParamInput		,   32, oMaster.FshoplinkerMallID) _
			,Array("@retErrStr"				, adVarchar		, adParamOutput		,  100, "") _
			,Array("@overseasPrice"			, adCurrency	, adParamInput		,     , oMaster.FoverseasPrice) _
			,Array("@overseasDeliveryPrice"	, adCurrency	, adParamInput		,     , oMaster.FoverseasDeliveryPrice) _
			,Array("@overseasRealPrice"		, adCurrency	, adParamInput		,     , oMaster.FoverseasRealPrice) _
			,Array("@reserve01"				, adVarchar		, adParamInput		,   32, oMaster.Freserve01) _
			,Array("@beasongNum11st"		, adVarchar		, adParamInput		,   16, oMaster.FbeasongNum11st) _
			,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FOutMallItemOption)) _
    	)

		if (IS_TEST_MODE = True) then
			sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_TEST"
'			response.write oMaster.FSellSite & "<br />"
'			response.write oMaster.FOutMallOrderSerial & "<br />"
'			response.write oMaster.FSellDate & "<br />"
'			response.write oMaster.FPayType & "<br />"
'			response.write oMaster.FPaydate & "<br />"
'			response.write oDetailArr(i).FItemID & "<br />"
'			response.write oDetailArr(i).FItemOption & "<br />"
'			response.write oDetailArr(i).FItemID & "<br />"
'			response.write oDetailArr(i).FOutMallItemName & "<br />"
'			response.write oDetailArr(i).FItemOption & "<br />"
'			response.write oDetailArr(i).FOutMallItemOptionName & "<br />"
'			response.write oMaster.FUserID & "<br />"
'			response.write oMaster.FOrderName & "<br />"
'			response.write oMaster.FOrderEmail & "<br />"
'			response.write oMaster.FOrderTelNo & "<br />"
'			response.write oMaster.FOrderHpNo & "<br />"
'			response.write oMaster.FReceiveName & "<br />"
'			response.write oMaster.FReceiveTelNo & "<br />"
'			response.write oMaster.FReceiveHpNo & "<br />"
'			response.write oMaster.FReceiveZipCode & "<br />"
'			response.write oMaster.FReceiveAddr1 & "<br />"
'			response.write oMaster.FReceiveAddr2 & "<br />"
'			response.write oDetailArr(i).Fitemcost & "<br />"
'			response.write oDetailArr(i).FReducedPrice & "<br />"
'			response.write oDetailArr(i).FItemNo & "<br />"
'			response.write oDetailArr(i).FdetailSeq & "<br />"
'			response.write oMaster.Fdeliverymemo & "<br />"
'			response.write oDetailArr(i).FrequireDetail & "<br />"
'			response.write oMaster.FdeliverPay & "<br />"
'			response.write oMaster.ForderCsGbn & "<br />"
'			response.write oMaster.FcountryCode & "<br />"
'			response.write oDetailArr(i).FOutMallItemID & "<br />"
'			response.write oMaster.Fshoplinkermallname & "<br />"
'			response.write oDetailArr(i).FshoplinkerPrdCode & "<br />"
'			response.write oMaster.FshoplinkerOrderID & "<br />"
'			response.write oMaster.FshoplinkerMallID & "<br />"
'			response.write oMaster.FoverseasPrice & "<br />"
'			response.write oMaster.FoverseasDeliveryPrice & "<br />"
'			response.write oMaster.FoverseasRealPrice & "<br />"
'			response.write oMaster.Freserve01 & "<br />"
'			response.write oMaster.FbeasongNum11st & "<br />"

			''dbget.rollbackTrans
			''dbget.close() : response.end
		else
			sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
		end if

			' If session("ssBctID")="kjy8517" Then
			' 	On Error Resume Next
			' 	dbget.BeginTrans
			' 	retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
			' 	If Err.Number <> 0 Then
			' 		tmpStr = " exec db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert "
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FSellSite) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOutMallOrderSerial) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FSellDate) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FPayType) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FPaydate) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemOption) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemOption) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemOptionName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FUserID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderEmail) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderTelNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderHpNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveTelNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveHpNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveZipCode) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveAddr1) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveAddr2) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).Fitemcost) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FReducedPrice) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FdetailSeq) & "',"
			' 		tmpStr = tmpStr + "'0',"
			' 		tmpStr = tmpStr + "'0',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.Fdeliverymemo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FrequireDetail) & "',"
			' 		tmpStr = tmpStr + "'" & orderDlvPay & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.ForderCsGbn & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FcountryCode & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemID) & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.Fshoplinkermallname & "',"
			' 		tmpStr = tmpStr + "'" & oDetailArr(i).FshoplinkerPrdCode & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FshoplinkerOrderID & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FshoplinkerMallID & "',"
			' 		tmpStr = tmpStr + "'',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasDeliveryPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasRealPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.Freserve01 & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FbeasongNum11st & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemOption) & "'"
			' 		rw tmpStr
			' 		rw "-----------------------------"
			' 	End If


			' 	RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
			' 	retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

			' 	if (RetErr<0) and (RetErr<>-1) then ''Break
			' 		'// 에러코드 -1 은 중복입력
			' 		dbget.rollbackTrans
			' 		if IsAutoScript then
			' 			response.write "ERROR["&retErr&"]"& retErrStr
			' 		else
			' 			response.write "ERROR["&retErr&"]"& retErrStr
			' 			response.write "<script>alert('오류가 발생했습니다.');</script>"
			' 		end if

			' 		dbget.close() : response.end
			' 	elseif (RetErr <> -1) then
			' 		SaveOrderToDB = True
			' 	end if

			' 	dbget.CommitTrans
			' 	On Error Goto 0
			' Else
			' 		dbget.BeginTrans

			' 		retParamInfo = fnExecSPOutput(sqlStr, paramInfo)

			' 		RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
			' 		retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

			' 		if (RetErr<0) and (RetErr<>-1) then ''Break
			' 			'// 에러코드 -1 은 중복입력
			' 			dbget.rollbackTrans
			' 			if IsAutoScript then
			' 				response.write "ERROR["&retErr&"]"& retErrStr
			' 			else
			' 				response.write "ERROR["&retErr&"]"& retErrStr
			' 				response.write "<script>alert('오류가 발생했습니다.');</script>"
			' 			end if

			' 			dbget.close() : response.end
			' 		elseif (RetErr <> -1) then
			' 			SaveOrderToDB = True
			' 		end if

			' 		dbget.CommitTrans
			' End If

		dbget.BeginTrans

		retParamInfo = fnExecSPOutput(sqlStr, paramInfo)

		RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
		retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

		if (RetErr<0) and (RetErr<>-1) then ''Break
			'// 에러코드 -1 은 중복입력
			dbget.rollbackTrans
			if IsAutoScript then
				response.write "ERROR["&retErr&"]"& retErrStr
			else
				response.write "ERROR["&retErr&"]"& retErrStr
				response.write "<script>alert('오류가 발생했습니다.');</script>"
			end if

			dbget.close() : response.end
		elseif (RetErr <> -1) then
			SaveOrderToDB = True
		end if

		dbget.CommitTrans
	next
end function

Function getCurrDateTimeFormat()
	Dim nowtimer : nowtimer= timer()
	getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	If NOT objfile.FolderExists(sFolderPath) Then
		objfile.CreateFolder sFolderPath
	End If
	Set objfile = Nothing
End Sub

Function DelAPITMPFile(iFileURI)
	Dim iFullPath
	iFullPath = server.mappath(replace(iFileURI,"http://wapi.10x10.co.kr",""))

	Dim FSO, iFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
		Set iFile = FSO.GetFile(iFullPath)
			If (iFile <> "") Then iFile.Delete
		Set iFile = Nothing
	Set FSO = Nothing
End Function

public function RequestArrayToArray(reqObj)
	dim obj, objArr()
	dim i
	Set obj = reqObj

	ReDim objArr(obj.Count - 1)

	For i = 0 To obj.Count - 1
		objArr(i) = obj(i+1)
	Next

	RequestArrayToArray = objArr
end function

function RemovePrecedingZero(str)
	dim result
	result = str
	do while (Left(result, 1) = "0")
		result = Mid(result, 2, 1000)
	loop
	RemovePrecedingZero = result
end function

Public Function getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

Public Function generateKey_nvstorefarm(iTimestamp)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		Else
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

'// 주문 발주처리
function PlaceProductOrder_nvstorefarm(ProductOrderID, isellsite)
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL
	dim strRst, objXML, xmlDOM

	PlaceProductOrder_nvstorefarm = False

	iServ		= "SellerService41"
	iCcd		= "PlaceProductOrder"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If isellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf isellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:PlaceProductOrderRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:ProductOrderID>"&ProductOrderID&"</sel:ProductOrderID>"
	strRst = strRst & "		</sel:PlaceProductOrderRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	PlaceProductOrder_nvstorefarm = True

	''set objMasterListXML = Nothing
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

function GetOrderDetailList_nvstorefarm(selldate, LastChangedStatusCode, isellsite)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL
	dim strRst, objXML, xmlDOM
	dim objMasterListXML, objMasterOneXML
	dim PrdOrderList(), i
	dim tmpXml

	Dim testStr1, testStr2
	testStr1 = request("testStr1")
	testStr2 = request("testStr2")

	redim PrdOrderList(-1)
	GetOrderDetailList_nvstorefarm = PrdOrderList

	iServ		= "SellerService41"
	iCcd		= "GetChangedProductOrderList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
If testStr1 <> "" Then
	strRst = strRst & "			<sel:InquiryTimeFrom>"& testStr1 &"</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"& testStr2 &"</sel:InquiryTimeTo>"										'조회 종료 일시(해당 시각 포함하지 않음)
Else
	strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
End If
	strRst = strRst & "			<sel:LastChangedStatusCode>" & LastChangedStatusCode & "</sel:LastChangedStatusCode>"				'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환 | PAYED : 결제완료)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'판매자 아이디
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"
If session("ssBctID")="kjy8517" Then
	rw objXML.responseText & "<br /><br />"
	rw "==================="
End If
	''dbget.close : response.end

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) = 0 then
		''if IsAutoScript then
			response.write "내역없음<br />"
		''end if

		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")

	i = 0
	redim PrdOrderList(objMasterListXML.length - 1)
	For each objMasterOneXML in objMasterListXML
		PrdOrderList(i) = objMasterOneXML.getElementsByTagName("n:ProductOrderID")(0).Text
		i = i + 1
	next

	GetOrderDetailList_nvstorefarm = PrdOrderList

	set objMasterListXML = Nothing
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

function GetOrderFrom_nvstorefarm(isellsite, selldate, chgCode)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	End If

	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim PrdOrderList, PrdOrder
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim cryptoLib
	dim keyGenerated
	Dim strSql, isDisCountYn, maySellPrice


	GetOrderFrom_nvstorefarm = False

	PrdOrderList = GetOrderDetailList_nvstorefarm(selldate, chgCode, sellsite)

	response.write "건수(" & UBound(PrdOrderList) + 1 & ") " & "<br />"

	if UBound(PrdOrderList) < 0 then
		exit function
	end if

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
	For each PrdOrder in PrdOrderList
		strRst = strRst & "			<sel:ProductOrderIDList>" & PrdOrder & "</sel:ProductOrderIDList>" + vbCrLf
	next
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
If session("ssBctID")="kjy8517" Then
	response.write objXML.responseText & "<br /><br />"
'	response.end
End If
	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (UBound(PrdOrderList) + 1) then
		response.write "건수 불일치 오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	keyGenerated = generateKey_nvstorefarm(iTimestamp)
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
	set objMasterListXML = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")
	i = 0
	For each objMasterOneXML in objMasterListXML

		if objMasterOneXML.getElementsByTagName("n:CancelInfo").length > 0 then
			'// 취소주문
		elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length < 1) then
			'// 주소입력 안된 주문(선물하기 주문은 받는 사람이 주소를 입력한 이후에 끌어와야 한다.)
		else
			Set oMaster = new COrderMasterItem
			isDisCountYn = "N"
			maySellPrice = ""
			oMaster.FSellSite 			= sellsite
			oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("n:Order/n:OrderID")(0).Text
			If oMaster.FOutMallOrderSerial = "2020121995761581" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2020121995761581_1"
			End If
			oMaster.FSellDate 			= Left(Now(), 10)
			oMaster.FPayType			= "50"
			oMaster.FPaydate			= oMaster.FSellDate
			oMaster.FOrderUserID		= ""
			oMaster.FOrderName			= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererName")(0).Text)), 28)
			if (objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2").length > 0) then
				oMaster.FOrderTelNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2")(0).Text))
			else
				oMaster.FOrderTelNo = html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			end if
			oMaster.FOrderHpNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			oMaster.FOrderEmail			= ""
			''response.Write objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length
			''response.end
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name")(0).Text)), 28)
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name")(0).Text)), 28)
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if
			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2")(0).Text))
			elseif objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2")(0).Text))
			else
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if
			end if

			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo").length > 0 then
				oMaster.Fdeliverymemo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo")(0).Text), 180)
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount").length > 0 then
				oMaster.FdeliverPay = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount")(0).Text
			end if

			If sellsite <> "nvstorefarmclass" Then
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode")(0).Text)
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode")(0).Text)
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress")(0).Text))
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress").length > 0) then
					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress")(0).Text))
				else
					oMaster.FReceiveAddr2		= "" '아래 주석 부분으로 했더니 출고지 주소가 출력 됨 (내용 -> 도봉동 여인닷컴)
'				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress").length > 0) then
'					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress")(0).Text))
				end if
				if InStr(oMaster.FReceiveZipCode, "-") = 0 then
					oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
				end if

				'// 주소 수정
				oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
				oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
				tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
				pos = 0
				for k = 0 to 2
					pos = InStr(pos+1, tmpStr, " ")
					if (pos = 0) then
						exit for
					end if
				next

				if (pos > 0) then
					oMaster.FReceiveAddr1 = Left(tmpStr, pos)
					oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
				end if

				oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
				oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)
			End If

			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOrderID")(0).Text
			oDetailArr(0).FItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode").length > 0) then
				oDetailArr(0).FItemOption = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode")(0).Text
			else
				oDetailArr(0).FItemOption = "0000"
			end if

			oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductID")(0).Text
			oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
			oDetailArr(0).FOutMallItemName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductName")(0).Text)
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption").length > 0) then
				oDetailArr(0).FOutMallItemOptionName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption")(0).Text)
			else
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:Quantity")(0).Text)

			'2019-08-06 김진영 아래 조건 추가
			'스토어팜 매입이면서 할인기간이라면 판매가(itemcost)를 실판매가(reducedprice)와 동일하게 저장
			'If left(now(),10) >= "2019-10-2" and left(now(),10) < "2019-09-24" Then
			'2019-10-21 김진영, 위 now()에서 Date로 변경 / Case SellerProductCode CSTR문자 변환, Trim 처리
			'2020-09-10 김진영, 스토어팜 특가관리에 추가했다면 할인가격으로 변경되게 수정
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
			strSql = strSql & " WHERE mallgubun = '"& sellsite &"' "
			strSql = strSql & " and itemid = '"& CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text)) &"' "
			strSql = strSql & " and GETDATE() >= startDate and GETDATE() <= endDate "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") > 0 Then
					isDisCountYn = "Y"
				Else
					isDisCountYn = "N"
				End If
			rsget.Close

			If isDisCountYn = "Y" Then
'				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 시작
				maySellPrice = Clng(objMasterOneXML.getElementsByTagName("n:UnitPrice")(0).Text)
				If (objMasterOneXML.getElementsByTagName("n:ProductImediateDiscountAmount").length > 0) then
					oDetailArr(0).Fitemcost = maySellPrice - Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductImediateDiscountAmount")(0).Text) / oDetailArr(0).FItemNo)
				Else
					oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
				End If
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 끝
			Else
				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			End If

			oDetailArr(0).FReducedPrice = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0


			if (SaveOrderToDB(oMaster, oDetailArr) = True) then
				if PlaceProductOrder_nvstorefarm(oDetailArr(0).FdetailSeq, sellsite) then
					successCnt = successCnt + 1
				end if
			end if
			i = i + 1
		end if
	next
	Set cryptoLib = Nothing

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	GetOrderFrom_nvstorefarm = True
	Set xmlDOM = Nothing
	Set objXML = Nothing

end function


Function GetOrder_nvstorefarm(isellsite, currdate, hasMoreData, chgCode, lastOrderNo, lastTime)
	Dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL, strSql
	dim strRst, objXML, xmlDOM
	dim objMasterListXML, objMasterOneXML, i

	iServ		= "SellerService41"
	iCcd		= "GetChangedProductOrderList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	If hasMoreData = "Y" Then
		strRst = strRst & "			<sel:InquiryTimeFrom>"&lastTime&"</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
		strRst = strRst & "			<sel:InquiryExtraData>"&lastOrderNo&"</sel:InquiryExtraData>"
	Else
		strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
	End If
	strRst = strRst & "			<sel:LastChangedStatusCode>" & chgCode & "</sel:LastChangedStatusCode>"								'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환 | PAYED : 결제완료)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'판매자 아이디
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML(objXML.responseText)
				' If session("ssBctID")="kjy8517" Then
				' 	rw objXML.responseText & "<br /><br />"
				' 	rw "==================="
				' End If
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					If xmlDOM.getElementsByTagName("n:HasMoreData").item(0).text = "true" Then
						hasMoreData = "Y"
						lastOrderNo	= xmlDOM.getElementsByTagName("n:InquiryExtraData").item(0).text
						lastTime	= xmlDOM.getElementsByTagName("n:MoreDataTimeFrom").item(0).text
					Else
						hasMoreData = "N"
					End If

					If CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) > 0 Then
						Set objMasterListXML = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")
							For Each objMasterOneXML in objMasterListXML
								strSql = ""
								strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] ([sellsite], [OutMallOrderSerial], [regdate]) "
								strSql = strSql & " VALUES ('"& sellsite &"', '"& objMasterOneXML.getElementsByTagName("n:ProductOrderID")(0).Text &"', '"& selldate &"') "
								dbget.Execute strSql
							Next
						Set objMasterListXML = nothing
					Else
						''if IsAutoScript then
							response.write "내역없음<br />"
						''end if

						Set xmlDOM = Nothing
						Set objXML = Nothing
						exit function
					End If
				Else
					response.write "오류 : 종료"
					Set xmlDOM = Nothing
					Set objXML = Nothing
					dbget.close : response.end
				End If
			Set xmlDOM = Nothing
		Else
			If IsAutoScript then
				response.write "ERROR : 통신오류"
			Else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			End If
			dbget.close : response.end
		End If
	Set objXML = Nothing
End Function

Function GetOrderFrom_NewCall_nvstorefarm(isellsite, currdate)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	End If
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst, arrRows
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim cryptoLib
	dim keyGenerated
	Dim strSql, isDisCountYn, maySellPrice, mayCnt
	GetOrderFrom_NewCall_nvstorefarm = False

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] "
	strSql = strSql & " WHERE sellsite = '"& sellsite &"' "
	strSql = strSql & " and regdate = '"& currdate &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		mayCnt = rsget("cnt")
	rsget.Close
	response.write "건수(" & mayCnt & "<br />"

	If mayCnt = 0 Then
		exit function
	End If

	strSql = ""
	strSql = strSql & " SELECT outmallOrderSerial "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] "
	strSql = strSql & " WHERE sellsite = '"& sellsite &"' "
	strSql = strSql & " and regdate = '"& currdate &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
	For i = 0 To ubound(arrRows,2)
		strRst = strRst & "			<sel:ProductOrderIDList>" & arrRows(0,i) & "</sel:ProductOrderIDList>" + vbCrLf
	next
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	' response.write strRst
	' dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
If session("ssBctID")="kjy8517" Then
	response.write objXML.responseText & "<br /><br />"
	response.end
End If
	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (mayCnt) then
		response.write "건수 불일치 오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	keyGenerated = generateKey_nvstorefarm(iTimestamp)
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
	set objMasterListXML = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")
	i = 0
	For each objMasterOneXML in objMasterListXML

		if objMasterOneXML.getElementsByTagName("n:CancelInfo").length > 0 then
			'// 취소주문
		elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length < 1) then
			'// 주소입력 안된 주문(선물하기 주문은 받는 사람이 주소를 입력한 이후에 끌어와야 한다.)
		else
			Set oMaster = new COrderMasterItem
			isDisCountYn = "N"
			maySellPrice = ""
			oMaster.FSellSite 			= sellsite
			oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("n:Order/n:OrderID")(0).Text
			If oMaster.FOutMallOrderSerial = "2020121995761581" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2020121995761581_1"
			End If
			oMaster.FSellDate 			= Left(Now(), 10)
			oMaster.FPayType			= "50"
			oMaster.FPaydate			= oMaster.FSellDate
			oMaster.FOrderUserID		= ""
			oMaster.FOrderName			= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererName")(0).Text)), 28)
			if (objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2").length > 0) then
				oMaster.FOrderTelNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2")(0).Text))
			else
				oMaster.FOrderTelNo = html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			end if
			oMaster.FOrderHpNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			oMaster.FOrderEmail			= ""
			''response.Write objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length
			''response.end
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name")(0).Text)), 28)
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name")(0).Text)), 28)
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if
			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2")(0).Text))
			elseif objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2")(0).Text))
			else
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if
			end if

			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo").length > 0 then
				oMaster.Fdeliverymemo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo")(0).Text), 180)
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount").length > 0 then
				oMaster.FdeliverPay = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount")(0).Text
			end if

			If sellsite <> "nvstorefarmclass" Then
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode")(0).Text)
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode")(0).Text)
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress")(0).Text))
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress").length > 0) then
					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress")(0).Text))
				else
					oMaster.FReceiveAddr2		= "" '아래 주석 부분으로 했더니 출고지 주소가 출력 됨 (내용 -> 도봉동 여인닷컴)
'				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress").length > 0) then
'					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress")(0).Text))
				end if
				if InStr(oMaster.FReceiveZipCode, "-") = 0 then
					oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
				end if

				'// 주소 수정
				oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
				oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
				tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
				pos = 0
				for k = 0 to 2
					pos = InStr(pos+1, tmpStr, " ")
					if (pos = 0) then
						exit for
					end if
				next

				if (pos > 0) then
					oMaster.FReceiveAddr1 = Left(tmpStr, pos)
					oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
				end if

				oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
				oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)
			End If

			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOrderID")(0).Text
			oDetailArr(0).FItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode").length > 0) then
				oDetailArr(0).FItemOption = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode")(0).Text
			else
				oDetailArr(0).FItemOption = "0000"
			end if

			oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductID")(0).Text
			oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
			oDetailArr(0).FOutMallItemName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductName")(0).Text)
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption").length > 0) then
				oDetailArr(0).FOutMallItemOptionName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption")(0).Text)
			else
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:Quantity")(0).Text)

			'2019-08-06 김진영 아래 조건 추가
			'스토어팜 매입이면서 할인기간이라면 판매가(itemcost)를 실판매가(reducedprice)와 동일하게 저장
			'If left(now(),10) >= "2019-10-2" and left(now(),10) < "2019-09-24" Then
			'2019-10-21 김진영, 위 now()에서 Date로 변경 / Case SellerProductCode CSTR문자 변환, Trim 처리
			'2020-09-10 김진영, 스토어팜 특가관리에 추가했다면 할인가격으로 변경되게 수정
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
			strSql = strSql & " WHERE mallgubun = '"& sellsite &"' "
			strSql = strSql & " and itemid = '"& CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text)) &"' "
			strSql = strSql & " and GETDATE() >= startDate and GETDATE() <= endDate "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") > 0 Then
					isDisCountYn = "Y"
				Else
					isDisCountYn = "N"
				End If
			rsget.Close

			If isDisCountYn = "Y" Then
'				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 시작
				maySellPrice = Clng(objMasterOneXML.getElementsByTagName("n:UnitPrice")(0).Text)
				If (objMasterOneXML.getElementsByTagName("n:ProductImediateDiscountAmount").length > 0) then
					oDetailArr(0).Fitemcost = maySellPrice - Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductImediateDiscountAmount")(0).Text) / oDetailArr(0).FItemNo)
				Else
					oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
				End If
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 끝
			Else
				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			End If

			oDetailArr(0).FReducedPrice = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0


			if (SaveOrderToDB(oMaster, oDetailArr) = True) then
				if PlaceProductOrder_nvstorefarm(oDetailArr(0).FdetailSeq, sellsite) then
					successCnt = successCnt + 1
				end if
			end if
			i = i + 1
		end if
	next
	Set cryptoLib = Nothing

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	GetOrderFrom_NewCall_nvstorefarm = True
	Set xmlDOM = Nothing
	Set objXML = Nothing
End Function

function GetCheckStatus(byVal sellsite, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp](sellsite, lastcheckdate, issuccess) "
	strSql = strSql + "		values('" & sellsite & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N') "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function GetCheckItemOptionValid(byVal itemid, byVal itemoption)
	dim strSql

	GetCheckItemOptionValid = False

    strSql = " select top 1 i.itemid "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    strSql = strSql + " 	and o.itemoption = '" & itemoption & "' "

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetCheckItemOptionValid = True
	end if
	rsget.Close
end function

function GetItemOptionWithOptionName(byVal sellsite, byVal itemid, byVal itemoptionname)
	dim strSql, found

	found = False
	GetItemOptionWithOptionName = "0000"



	'// 모델명:SMN-204 you're in
	itemoptionname = Replace(itemoptionname, "'", "''")


	if (sellsite = "ezwel") then
		strSql = "exec [db_temp].[dbo].[usp_TEN_xSiteOrder_OptionMapping_EzWel] '"&itemid&"','"&itemoptionname&"'"

		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			GetItemOptionWithOptionName = rsget("itemoption")
			found = True
		end if
		rsget.Close
	end if

	if found then
		exit function
	end if

    strSql = " select top 1 o.itemoption "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    strSql = strSql + " 	and o.optionname = '" & itemoptionname & "' "
    ''response.Write strSql

	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetItemOptionWithOptionName = rsget("itemoption")
		found = True
	end if
	rsget.Close

	if found then
		exit function
	end if

	'사방넷을 통한 LFmall 옵션명에  None[XX]: 글자가 심어서 넘어옴
	If Instr(itemoptionname, "None[XX]:") > 0 Then
		itemoptionname = Replace(itemoptionname, "None[XX]:", "")
	End If

    strSql = " select top 1 o.itemoption "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    ''strSql = strSql + " 	and (o.optionname = '" & Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/") & "') "
    strSql = strSql + " 	and (Replace(Replace(o.optionname, ',', ''), ':', '') = '" & Replace(Replace(Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/"), ",", ""), ":", "") & "') "

	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetItemOptionWithOptionName = rsget("itemoption")
		found = True
	end if
	rsget.Close

	if found then
		exit function
	end if

	if (sellsite = "lotteCom") and False then
	    strSql = " select top 1 o.itemoption "
	    strSql = strSql + " from "
	    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
	    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
	    strSql = strSql + " 	on "
	    strSql = strSql + " 		i.itemid = o.itemid "
	    strSql = strSql + " where "
	    strSql = strSql + " 	1 = 1 "
	    strSql = strSql + " 	and i.itemid = " & itemid
	    strSql = strSql + " 	and (Replace(Replace(o.optionname, ',', ''), ':', '') = '" & Replace(Replace(Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/"), ",", ""), ":", "") & "') "

		rsget.CursorLocation = adUseClient
    	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			GetItemOptionWithOptionName = rsget("itemoption")
			found = True
		end if
		rsget.Close

		if found then
			exit function
		end if
	end if

end function

function SetCheckStatus(sellsite, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "' "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

function arrayMerge(left, right)
	dim right_size
	dim total_size
	dim i
	dim merged
	''// Convert "left" to an array
	if not isArray(left) then
		left = Array(left)
	end if
	''// Convert "right" to an array
	if not isArray(right) then
		right = Array(right)
	end if
	''// Start with "left" and add the elements of "right"

	right_size = ubound(right)
	total_size = ubound(left) + right_size + 1

	merged = array()
	redim merged(total_size)
	dim counter : counter = 0

	for i = lbound(left) to ubound(left)
		if isobject(left(i))then
			set merged(counter) = left(i)
		else
			merged(counter) = left(i)
		end if
		counter=counter+1
	next

	for i = lbound(right) to ubound(right)
		if isobject(right(i))then
			set merged(counter) = right(i)
		else
			merged(counter) = right(i)
		end if
	next


	''// Return value
	arrayMerge = merged
end function


public function getDelimCharCount(orgStr, delim)
    dim retCNT : retCNT = 0
    dim buf
    buf = split(orgStr,delim)

    if IsArray(buf) then
        retCNT = UBound(buf)
    end if
    getDelimCharCount = retCNT
end function

'' SSG 매칭된 옵션코드 리턴 [빨강/XL/소재/3] [화이트] [빨강/L,1:^:주문제작문구:^:문구작성] [,1:^:주문제작문구:^:주문문구1,2:^:주문제작문구:^:주문문구2]  '' spliter ,  / :^:
function getOptionCodByOptionNameSSG(iitemid,outmalloptionName,byref requiredtl)
    dim retStr, sqlStr : retStr=""
    dim ichrCnt, IsDoubleOption, IsTreepleOption
    dim ioptionname, ireqdrlname

    if (outmalloptionName="") then
        requiredtl = ""
        getOptionCodByOptionNameSSG = "0000"
        Exit function
    end if

    ioptionname = outmalloptionName
    ichrCnt = getDelimCharCount(ioptionname,",")

''////////////////////////////////////////////////////// 예전 버전 ////////////////////////////////////////////////////////
'     IF (ichrCnt>=1) THEN ''주문제작 문구가 있는 상품
'         ioptionname = split(outmalloptionName,",")(0)
'         requiredtl  = replace(split(outmalloptionName,",")(1),"1:^:주문제작문구:^:","")
'         ''requiredtl  = replace(split(outmalloptionName,",")(1),"1:^:asdasd:^:","")

'         if ichrCnt>1 then
'             requiredtl = requiredtl + ","+replace(split(outmalloptionName,",")(2),"2:^:주문제작문구:^:","")
'             ''requiredtl = requiredtl + ","+replace(split(outmalloptionName,",")(2),"2:^:asdasdddd:^:","")
'         end if
'   ''rw "[requiredtl]"&requiredtl
'         'rw "ioptionname:"&ioptionname
'         'rw "requiredtl:"&requiredtl
'     end if
''////////////////////////// 수정 버전 2019-12-11 11:40 김진영 수정 주문번호 :19120982908 문제 발생  ////////////////////////////
    IF (ichrCnt>=1) THEN ''주문제작 문구가 있는 상품
        ioptionname = split(outmalloptionName,",")(0)
		If instr(outmalloptionName, "1:^:주문제작문구:^:") > 0 Then
			requiredtl = Split(outmalloptionName, "1:^:주문제작문구:^:")(1)
			If instr(requiredtl, "2:^:주문제작문구:^:") > 0 then
				requiredtl = Replace(requiredtl, "2:^:주문제작문구:^:", "")
			end if
		End If
    end if
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if (ioptionname="") then  ''주문제작문구를 발라낸후 옵션명이 없으면.
        getOptionCodByOptionNameSSG = "0000"
        Exit function
    end if

    IF (getDelimCharCount(ioptionname,"/")=1) THEN
        IsDoubleOption = TRUE
    ELSEIF (getDelimCharCount(ioptionname,"/")=2) THEN  '''빨강/XL/소재/3 = 옵션명에 / 가 있을경우 못발라낼 수 있음.(소재/3)
        IsTreepleOption = TRUE
    ENd IF


    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and optionname='"&replace(ioptionname,"/",",")&"'"   ''replace(optionname,'*','')
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and optionname='"&replace(ioptionname,"/",",")&"'"   ''replace(optionname,'*','')
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and Replace(optionname,',','')='"&ioptionname&"'"
    END IF

''response.write sqlstr & "<Br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

	''옵션명에 "/" 가 있는 CASE===============================================================================
	If (retStr="") THEN
		sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteOrder_OptionMapping_SSG] '"&iitemid&"','"&replace(Trim(outmalloptionName),"'","''")&"'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			retStr = rsget("itemoption")
		end if
		rsget.Close
	END IF
	''=====================================================================================================

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행 ?  0000 맞나?
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>0) THEN
    	        retStr = "FF00" '"0000"=>FF00
    	    else
    	        retStr = "0000"
    	    end if
    	end if
        rsget.Close
    END IF

    getOptionCodByOptionNameSSG = retStr
end function

function RemoveWhiteSpaceChar(str)
	dim retVal
	If isNull(str) Then
		RemoveWhiteSpaceChar = ""
		Exit Function
	End If

	retVal = str
	retVal = Replace(retVal, Chr(13), "")
	retVal = Replace(retVal, Chr(10), "")
	retVal = Replace(retVal, vbTab, " ")
	retVal = Trim(retVal)
	RemoveWhiteSpaceChar = retVal
end function

Function getApiUrl(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiUrl = "https://dev-openapi.lotteon.com"
			Else
				getApiUrl = "https://openapi.lotteon.com"
			End If
	End Select
End Function

Function getApiKey(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiKey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
			Else
				getApiKey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
			End If
	End Select
End Function
%>
