<%
Dim GG_IS_DEVSERVER : GG_IS_DEVSERVER = (application("Svr_Info") = "Dev")
Dim GG_CACHEDB_APP_CON_NAME : GG_CACHEDB_APP_CON_NAME = "db_cache"

Dim GG_SPLIT_MEM_ARRAY  : GG_SPLIT_MEM_ARRAY = "[@||@]"
Dim GG_SPLIT_MEM_KEYVAL : GG_SPLIT_MEM_KEYVAL = "[|@|]"
Dim GG_SPLIT_MEM_VALARRAY : GG_SPLIT_MEM_VALARRAY = "[@|@]"
Dim GG_MEM_KEY_ARRAY_ENDFIX : GG_MEM_KEY_ARRAY_ENDFIX = ":::" 



function deWrite(iMsg)
    if (Not GG_IS_DEVSERVER) then Exit function
    response.write iMsg
end function

function fnDBCacheHashKey(ipreFix,iKey)
    dim Obj
    Set Obj = Server.CreateObject("nonnoi_ASPEncrypt.ASPEncrypt")  ''object is faster then asp code
    Obj.RegisterName = "SEO SEOK"
    Obj.RegisterKey  = "63918C68A2D78AF7-5755"
    Obj.HashAlgorithm = 0 ''md5
    fnDBCacheHashKey = ipreFix&":"&Obj.HashString(iKey)
    set Obj = nothing
end function

function SerializeRs(iRs) ''return is Stream
    dim istream
    set istream = server.CreateObject("ADODB.Stream")
    istream.Type = 1
    iRs.Save istream,0
    
    set SerializeRs = istream
    set istream = nothing
end function

function DeserializeRs(iBArr) ''return is RecordSet
    dim istream
    set istream = server.CreateObject("ADODB.Stream")
    istream.Type = 1
    istream.Open
    istream.Write iBArr
    istream.Position = 0
    
    dim irs
    set irs = server.CreateObject("ADODB.Recordset")
    irs.open istream
    set DeserializeRs = irs
    
    set istream = Nothing
    set irs = Nothing
end function 

function SetDBCacheTxtVal(ikey,iVal,iQuery,icacheSec)
    dim iCacheCon : set iCacheCon = CreateObject("ADODB.Connection")
    Dim cmd : set cmd = server.CreateObject("ADODB.Command")
    Dim strSql
    
    strSql = "db_Cache.dbo.sp_ten_setVal_Txt"
    
    iCacheCon.Open Application(GG_CACHEDB_APP_CON_NAME) ''커넥션 스트링.
    cmd.ActiveConnection = iCacheCon
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ikey", adVarchar, adParamInput, 100, ikey)
    cmd.Parameters.Append cmd.CreateParameter("@ival", adVarChar, adParamInput, 25000 , iVal)
    cmd.Parameters.Append cmd.CreateParameter("@icacheTime", adInteger, adParamInput, , icacheSec)
    cmd.Parameters.Append cmd.CreateParameter("@iquery", adVarchar, adParamInput, 3000, iQuery)
    cmd.Execute 
    set cmd = Nothing
    iCacheCon.Close
    SET iCacheCon = Nothing
end function

function getDBCacheTxtVal(ikey,byref iVal)
    dim iCacheCon : set iCacheCon = CreateObject("ADODB.Connection")
    dim iCacheRs : set iCacheRs = CreateObject("ADODB.Recordset")
    dim strSql
 dim otime : otime=Timer()   
    getDBCacheTxtVal = false
    
    strSql = "[db_Cache].[dbo].sp_ten_getVal_TXT('"&ikey&"')"
    
    iCacheCon.Open Application(GG_CACHEDB_APP_CON_NAME) ''커넥션 스트링.
    iCacheRs.Open strSql, iCacheCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    
    if not iCacheRs.EOF then
        getDBCacheTxtVal = true
        iVal = iCacheRs(0)
        deWrite "hit!"&":"&FormatNumber(Timer()-otime,2)
    end if
    iCacheRs.Close
    iCacheCon.Close
    SET iCacheCon = Nothing
end function

function getDBCacheTxtValNoCheck(ikey,byref iVal)
    dim iCacheCon : set iCacheCon = CreateObject("ADODB.Connection")
    dim iCacheRs : set iCacheRs = CreateObject("ADODB.Recordset")
    dim strSql
 dim otime : otime=Timer()   
    getDBCacheTxtValNoCheck = false
    
    strSql = "[db_Cache].[dbo].sp_ten_getVal_TXT_Nocheck('"&ikey&"')"
    
    iCacheCon.Open Application(GG_CACHEDB_APP_CON_NAME) ''커넥션 스트링.
    iCacheRs.Open strSql, iCacheCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    
    if not iCacheRs.EOF then
        getDBCacheTxtValNoCheck = true
        iVal = iCacheRs(0)
        deWrite "hit!"&":"&FormatNumber(Timer()-otime,2)
    end if
    iCacheRs.Close
    iCacheCon.Close
    SET iCacheCon = Nothing
end function

function getDBCacheSQL(idbget,irsget,ipreFix,isql,icacheSec) ''return is RecordSet
    dim strSql, ikey
    dim iCacheCon : set iCacheCon = CreateObject("ADODB.Connection")
    dim iCacheRs : set iCacheRs = CreateObject("ADODB.Recordset")
    dim isCacheExists
    
    ''캐시 디비에서 검색.
    ikey = fnDBCacheHashKey(ipreFix,isql)
  'rw "ikey:"&ikey
    strSql = "[db_Cache].[dbo].sp_ten_getVal('"&ikey&"')"
    
    iCacheCon.Open Application(GG_CACHEDB_APP_CON_NAME) ''커넥션 스트링.
    iCacheRs.Open strSql, iCacheCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    
    if  not iCacheRs.EOF  then
        isCacheExists = true
        SET getDBCacheSQL = DeserializeRs(iCacheRs(0))
    end if
    iCacheRs.Close
    
    
    ''캐시 디비에 존재하면 리턴 후 종료
    if (isCacheExists) then 
        iCacheCon.Close
        set iCacheRs = Nothing
        set iCacheCon = Nothing
        deWrite "<font color=red>hit:</font>"
        Exit function
    end if
        
    
    irsget.CursorLocation = adUseClient
    irsget.Open isql, idbget, adOpenForwardOnly, adLockReadOnly
    
    dim iStream, iSize, ibuf
    set iStream = SerializeRs(irsget)
    
    iSize = iStream.Size
    ibuf = iStream.read  

    Dim cmd : set cmd = server.CreateObject("ADODB.Command")
    strSql = "db_Cache.dbo.sp_ten_setVal"
    
    cmd.ActiveConnection = iCacheCon
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ikey", adVarchar, adParamInput, 100, ikey)
    cmd.Parameters.Append cmd.CreateParameter("@ival", adVarBinary, adParamInput, iSize, ibuf)
    cmd.Parameters.Append cmd.CreateParameter("@icacheTime", adInteger, adParamInput, , icacheSec)
    cmd.Parameters.Append cmd.CreateParameter("@iquery", adVarchar, adParamInput, 3000, isql)
    cmd.Execute 
    set cmd = Nothing


    set getDBCacheSQL = irsget
    ''irsget.close
    ''set irsget = Nothing

    iCacheCon.Close
    set iCacheRs = Nothing
    set iCacheCon = Nothing
        
end function


''------------------------------------

''Array를 Memcashed용 String으로 변환
'function Arr2StrSerialize(iArr)
'    dim i, retVal
'    retVal = ""
'    for i=LBound(iArr) to UBound(iArr)
'        retVal = retVal & Trim(iArr(i)) 
'        if i<UBound(iArr) then
'            retVal = retVal & GG_SPLIT_MEM_VALARRAY
'        end if
'    next
'    Arr2StrSerialize = retVal
'end function
'
'''serialize Dictionary to String
'''key[||]value@||@key[||]value , array spliter-@|@
'function Dic2StrSerialize(iDic)
'    dim i,cnt, ikey, ival
'    dim j, iArrStr
'    cnt = iDic.Count
'    
'    ''rw "iDic.Count="&cnt
'    for i=0 to cnt-1
'        'rw "iDic.item(i)="&iDic.item(i)&TypeName(iDic.item(i))&(iDic.item(i))
'        'rw "iDic.items()(i)="&iDic.items()(i)&TypeName(iDic.items()(i))&(iDic.items()(i))
'        ikey = Trim(iDic.Keys()(i))
'        ival = iDic.item(ikey)
'        ''rw ikey&":"&TypeName(ival)&":"&isArray(ival)
'        
'        if isArray(ival) then
'            ival = Arr2StrSerialize(ival)
'            ikey = ikey & GG_MEM_KEY_ARRAY_ENDFIX  ''밸류가 Array 인경우 구분
'        end if
'        
'        iArrStr = iArrStr & ikey &GG_SPLIT_MEM_KEYVAL& Trim(ival) & GG_SPLIT_MEM_ARRAY
'    next
'    Dic2StrSerialize = iArrStr
'end Function
'
'function Str2DicUnSerialize(iStr,iDic)
'    dim i, iArr1, iKey, iVal
'    iDic.RemoveAll()
'    iArr1 = split(iStr,GG_SPLIT_MEM_ARRAY)
'    for i=LBound(iArr1) to UBound(iArr1)
'        if (iArr1(i)<>"") then
'            iKey = SplitValue(iArr1(i),GG_SPLIT_MEM_KEYVAL,0)
'            iVal = SplitValue(iArr1(i),GG_SPLIT_MEM_KEYVAL,1) 
'            
'            ''if (InStr(iVal,GG_SPLIT_MEM_VALARRAY)>0) then
'            ''rw "iKey=="&iKey&"//"
'            if (RIGHT(iKey,LEN(GG_MEM_KEY_ARRAY_ENDFIX))=GG_MEM_KEY_ARRAY_ENDFIX) then
'                iKey = LEFT(iKey,LEN(iKey)-LEN(GG_MEM_KEY_ARRAY_ENDFIX))
'                iVal = split(iVal,GG_SPLIT_MEM_VALARRAY)
'                ''rw "iKey="&iKey
'            end if
'            iDic.Item(iKey)=iVal
'        end if
'    Next
'end function
'
'
''' simple value get
'function MemValGetExists(ipreFix,iKeyStr,byref iVal)
'    dim ikey : ikey = fnDBCacheHashKey(ipreFix,iKeyStr)
'    MemValGetExists = getDBCacheTxtVal(ikey,ival)
'end function
'
''' simple value set
'function MemValSet(ipreFix,iKeyStr,iVal,icacheSec)
'    call SetDBCacheTxtVal(fnDBCacheHashKey(ipreFix,iKeyStr),iVal,iKeyStr,icacheSec)
'end function
'
''' dictionary value set
'function  MemValSetDic(ipreFix,iKeyStr,iDic,icacheSec)
'    dim valStr
'    if (LCASE(TypeName(iDic))="dictionary") then
'        valStr = Dic2StrSerialize(iDic)
'        call SetDBCacheTxtVal(fnDBCacheHashKey(ipreFix,iKeyStr),valStr,iKeyStr,icacheSec)
'    end if
'end function
'
''' DbCache-Docruzerd Wrapper
'Class CDocWrapper
'    private FDocruzer
'    private FuseMemcache
'    private FQueryType  '' 1:value(count),  2:Array
'    private FKeyPreFix
'    
'    private FOrignQuery
'    private FCacheSec
'    
'    private FMemExists
'    private FMemVal
'    private FMemArrayStr
'    private Fcollec
'    
'    public LC_KOREAN 
'    public CS_EUCKR  
'    
'    private function getDicValueMem(iarrKey)
'        getDicValueMem = Fcollec.Item(iarrKey)
'    end function
'    
'    private function setDicValueMem(iarrKey,iarrVal)
'        Fcollec.Item(iarrKey) = iarrVal
'        ''Fcollec.Add iarrKey,iarrVal
'        ''rw "iarrKey,iarrVal="&iarrKey&","&TypeName(iarrVal)
'    end function
'
'    public function BeginSession()
'        if (FuseMemcache) then
'            
'        else
'            BeginSession = FDocruzer.BeginSession()
'        end if
'    end function
'
'    public function SubmitQuery(iSvrAddr, iSvrPort, _
'						iAuthCode, iLogs, iScn, _
'						iSearchQuery,iSortQuery, _
'						iFRectSearchTxt,iStartNum, iFPageSize, _
'						iDocruzerLC, iDocruzerCS)
'	''dim otime : otime=Timer()
'        if (FuseMemcache) then
'            FOrignQuery = iSearchQuery&iSortQuery&iStartNum&iFPageSize
'            if (FQueryType=1) then
'                if (MemValGetExists(FKeyPreFix,FOrignQuery,FMemVal)) then 
'                    FMemExists = True
'                    SubmitQuery = 1
'                    EXIT Function
'                end if
'            elseif (FQueryType=2) then
'                if (MemValGetExists(FKeyPreFix,FOrignQuery,FMemArrayStr)) then 
'                    call Str2DicUnSerialize(FMemArrayStr,Fcollec)
'                    FMemExists = True
'                    SubmitQuery = 1
'                    EXIT Function
'                end if
'            end if
'        end if
'        
'        if (FuseMemcache) then  '' use memcached but Miss
'            SET FDocruzer = Server.CreateObject("ATLDocruzer_3_2.Client")
'            if (FDocruzer.BeginSession<0) then 
'                SubmitQuery = -1
'                Exit function 
'            end if
'        end if
'    
'        SubmitQuery = FDocruzer.SubmitQuery(iSvrAddr, iSvrPort, _
'					iAuthCode, iLogs, iScn, _
'					iSearchQuery,iSortQuery, _
'					iFRectSearchTxt,iStartNum, iFPageSize, _
'					iDocruzerLC, iDocruzerCS)
'	 ''debugWrite "doc"&":"&FormatNumber(Timer()-otime,4)
'    end function
'
'    public function GetResult_TotalCount(byRef iTotalCount)
'        if (FuseMemcache) then
'            if (FMemExists) and (FQueryType=1) then
'                iTotalCount = FMemVal
'                Exit function
'            end if
'            
'            if (FMemExists) and (FQueryType=2) then
'                iTotalCount = getDicValueMem("TotalCount")
'                Exit function
'            end if
'        end if
'        
'        Call FDocruzer.GetResult_TotalCount(iTotalCount)
'        
'        if (FuseMemcache) and (FQueryType=1) then ''멤캐시 사용하고 데이타 메모리에 존재 안하면
'            FMemVal = iTotalCount
'        end if
'        
'        if (FuseMemcache) and (FQueryType=2) then 
'            call setDicValueMem("TotalCount",iTotalCount)
'        end if
'        
'    end function
'
'    public function GetResult_RowSize(byRef iResultcount) ''검색결과수
'        if (FuseMemcache) then
'            if (FMemExists) and (FQueryType=2) then
'                iResultcount = getDicValueMem("RowSize")
'                Exit function
'            end if
'        end if
'        
'        call FDocruzer.GetResult_RowSize(iResultcount)
'        
'        if (FuseMemcache) and (FQueryType=2) then 
'            call setDicValueMem("RowSize",iResultcount)
'        end if
'    end function
'    
'    public function GetResult_Rowid(byRef iRowids,byRef iScores)
'        if (FuseMemcache) then
'            if (FMemExists) and (FQueryType=2) then
'                iRowids = getDicValueMem("Rowids")
'                iScores = getDicValueMem("Scores")
'                Exit function
'            end if
'        end if
'        
'        call FDocruzer.GetResult_Rowid(iRowids, iScores)
'        
'        if (FuseMemcache) and (FQueryType=2) then 
'            call setDicValueMem("Rowids",iRowids)
'            call setDicValueMem("Scores",iScores)
'        end if
'    end function
'    
'    public function GetResult_Row(byRef iarrData,byRef iarrSize, iRows )
'        if (FuseMemcache) then
'            if (FMemExists) and (FQueryType=2) then
'                iarrData = getDicValueMem("arrData"&iRows)
'                Exit function
'            end if
'        end if
'        
'        call FDocruzer.GetResult_Row(iarrData, iarrSize, iRows)
'        
'        if (FuseMemcache) and (FQueryType=2) then 
'            call setDicValueMem("arrData"&iRows,iarrData)
'        end if
'    end function
'    
'    public function EndSession()
'        if (Not FuseMemcache) then
'            FDocruzer.EndSession()
'        else
'            ''Store Value
'            if (Not FMemExists) then
'                if (FQueryType=1) then 
'                    Call MemValSet(FKeyPreFix,FOrignQuery,FMemVal,FCacheSec)
'                end if
'                
'                if (FQueryType=2) then 
'                    Call MemValSetDic(FKeyPreFix,FOrignQuery,Fcollec,FCacheSec)
'                end if
'            end if
'        end if
'    end function
'
'    public function InItWrapper(isUseMem,iKeyPreFix,iQueryType,icacheSec)
'        FuseMemcache = isUseMem
'        FQueryType   = iQueryType
'        FCacheSec    = icacheSec
'        FKeyPreFix   = iKeyPreFix
'        if (FuseMemcache) then
'            Set Fcollec = CreateObject("Scripting.Dictionary")
'            
'            'rw "IsObject(Fcollec)"&IsObject(Fcollec)
'            'rw "TypeName(Fcollec)"&TypeName(Fcollec)
'            'rw "IsArray(Fcollec)"&IsArray(Fcollec)
'        else
'            SET FDocruzer = Server.CreateObject("ATLDocruzer_3_2.Client")
'        end if
'    end function
'
'    Private Sub Class_Initialize()
'        FMemExists = false
'        LC_KOREAN = 1
'        CS_EUCKR  = 1
'	End Sub
'
'	Private Sub Class_Terminate()
'        SET FDocruzer = NOTHING
'        SET Fcollec   = NOTHING
'	End Sub
'	
'End Class
%>