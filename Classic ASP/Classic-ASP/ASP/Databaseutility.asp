<%
option explicit 
 '---- CursorTypeEnum Values ----
 Const adOpenForwardOnly = 0
 Const adOpenKeyset = 1
 Const adOpenDynamic = 2
 Const adOpenStatic = 3

 '---- LockTypeEnum Values ----
 Const adLockReadOnly = 1
 Const adLockPessimistic = 2
 Const adLockOptimistic = 3
 Const adLockBatchOptimistic = 4

 '------ CommandType Enum Values ---
  Const  adCmdUnspecified = -1
  Const  adCmdText = 1
  Const  adCmdTable = 2
  Const  adCmdStoredProc = 4
  Const  adCmdUnknown = 8
  Const  adCmdFile = 256
  Const  adCmdTableDirect = 512
 
  '------ CursorLocation Enum Values ---  
   Const  adUseClient = 3
   Const  adUseNone = 1
   Const  adUseServer = 2

   '------- dataType Constants --------
const adEmpty	            = 0	        'No value
const adSmallInt	        = 2	        'A 2-byte signed integer.
const adInteger	            = 3	        'A 4-byte signed integer.
const adSingle	            = 4	        'A single-precision floating-point value.
const adDouble	            = 5	        'A double-precision floating-point value.
const adCurrency	        = 6	        'A currency value
const adDate	            = 7	        'The number of days since December 30, 1899 + the fraction of a day.
const adBSTR	            = 8	        'A null-terminated character string.
const adIDispatch	        = 9	        'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
const adError	            = 10	    'A 32-bit error code
const adBoolean	            = 11	    'A boolean value.
const adVariant	            = 12	    'An Automation Variant. Note: Currently not supported by ADO.
const adIUnknown	        = 13	    'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
const adDecimal	            = 14	    'An exact numeric value with a fixed precision and scale.
const adTinyInt	            = 16	    'A 1-byte signed integer.
const adUnsignedTinyInt	    = 17	    'A 1-byte unsigned integer.
const adUnsignedSmallInt	= 18	    'A 2-byte unsigned integer.
const adUnsignedInt	        = 19	    'A 4-byte unsigned integer.
const adBigInt	            = 20	    'An 8-byte signed integer.
const adUnsignedBigInt	    = 21	    'An 8-byte unsigned integer.
const adFileTime	        = 64	    'The number of 100-nanosecond intervals since January 1,1601
const adGUID	            = 72	    'A globally unique identifier (GUID)
const adBinary	            = 128	    'A binary value.
const adChar	            = 129	    'A string value.
const adWChar	            = 130	    'A null-terminated Unicode character string.
const adNumeric	            = 131	    'An exact numeric value with a fixed precision and scale.
const adUserDefined	        = 132	    'A user-defined variable.
const adDBDate	            = 133	    'A date value (yyyymmdd).
const adDBTime	            = 134	    'A time value (hhmmss).
const adDBTimeStamp	        = 135	    'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
const adChapter	            = 136	    'A 4-byte chapter value that identifies rows in a child rowset
const adPropVariant	        = 138	    'An Automation PROPVARIANT.
const adVarNumeric	        = 139	    'A numeric value (Parameter object only).
const adVarChar	            = 200	    'A string value (Parameter object only).
const adLongVarChar	        = 201	    'A long string value.
const adVarWChar	        = 202	    'A null-terminated Unicode character string.
const adLongVarWChar	    = 203	    'A long null-terminated Unicode string value.
const adVarBinary	        = 204	    'A binary value (Parameter object only).
const adLongVarBinary	    = 205	    'A long binary value.



    '------- Parameter Direction Enums------------
    const adParamUnknown	    = 0	    'Direction is unknown
    const adParamInput	        = 1	    'Input parameter
    const adParamOutput	        = 2	    'Output parameter
    const adParamInputOutput	= 3	    'Both input and output parameter
    const adParamReturnValue	= 4	    'Return value

    '-----------ExecuteOptionEnum Enums -----------------
    Const adAsyncExecute = &H00000010
    Const adAsyncFetch = &H00000020
    Const adAsyncFetchNonBlocking = &H00000040
    Const adExecuteNoRecords = &H00000080
    Const adExecuteStream = &H00000400
 dim sConn

sConn = "Provider=SQLOLEDB.1;Data Source=W107VVLN53;Initial Catalog=Work;User Id=sa;password=Newuser@123;Persist Security Info=False;"

function ExecuteStoreProcedure(ByVal strSp, params, ByRef OutArray, ByVal OutPutParams)
    dim cmd , connection
    set cmd = Server.CreateObject("adodb.Command")
    set connection = Server.CreateObject("adodb.Connection")
    connection.Open sConn
     with cmd
            .ActiveConnection = connection
            .CommandText      = strSp
            .CommandType      = adCmdStoredProc 'Ado Command TypeEnum adCmdStoredProc
    end with

    parameterCollection cmd , params

    cmd.Execute , , adExecuteNoRecords
    if OutPutParams then OutArray = collectOutPutParams(cmd,params)

    set cmd.ActiveConnection = nothing
    set cmd = nothing
    ExecuteStoreProcedure = 0

end function

function GetRecordSet(ByVal strStoreProcName, params)

    dim cmd, OutputParams, ResultSet, connection
    set ResultSet = Server.CreateObject("adodb.Recordset")
    set cmd = Server.CreateObject("adodb.Command")

    set connection   = Server.CreateObject("adodb.Connection")
    connection.Open sConn

    with cmd
            .ActiveConnection = connection
            .CommandText      = strStoreProcName
            .CommandType      = adCmdStoredProc
    end with

    parameterCollection cmd , params

    ResultSet.CursorLocation    = adUseClient
    ResultSet.Open cmd, , adOpenStatic, adLockReadOnly 'adOpenStatic = 3 ' adLockReadOnly=1

    set cmd.ActiveConnection = nothing
    set cmd = nothing
    
    
    set GetRecordSet = ResultSet
end function

function CreateParameters(ByVal paramName, ByVal paramDataType, ByVal paramDirection, ByVal paramSize, ByVal paramValue)
    dim param
    set param = Server.CreateObject("adodb.Parameter")

    with param
              .Name = paramName
              .Type = paramDataType
    if not isEMpty(paramSize) then
             .Size  = paramSize
    end if
    if isEMpty(paramValue) then
             paramValue = null
    end if
            .Direction  =   paramDirection
            .Value      = paramValue
    end with

    set CreateParameters = param
end function

function ExecuteCommand(byVal sp_Name, byval ParameterArray, byRef strError)
    dim  output_Array
    Err.Clear

    on Error Resume next
    if not isEmpty(sp_Name) then
        call ExecuteStoreProcedure(sp_Name, ParameterArray, output_Array, 1)

        if Err.Number <> 0 then
            strError = Err.Description
            exit function
        end if

        if isArray(output_Array) then
            if not isEmpty(output_Array(0)) then
                ExecuteCommand = trim(output_Array(0))
               
            end if
        end if
    end if
    on error Goto 0
end function

function collectOutPutParams(byref cmd, argparams)
    dim params, v, OutArray(), Param

    dim i, l, u

    for each Param in argparams
        if param.direction = adParamOutPut then
            l = l+1
        end if
    next

    redim OutArray(l-1)
    u=0
    params = argparams
    for i= LBound(params) To UBound(params)
        if params(i).direction = adParamaOutput then
            OutArray(u) = cmd.Parameters(i).value
            u = u + 1
        end if
    next
    collectOutPutParams = OutArray

end function

sub parameterCollection(byRef cmd, Byval paramArrays)
    dim index
    with cmd
            if not isArray(paramArrays) then
                exit sub
            end if
        for index = LBound(paramArrays) To UBound(paramArrays)
                .Parameters.Append paramArrays(index)
        next
    end with
end sub
%>