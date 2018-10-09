<%
function getAutoQuantity(jobnumber,sheetnumber,jobtype)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
getAutoQuantity=""
set rsS=server.CreateObject("adodb.recordset")
if jobtype="N" and left(jobnumber,3)<>"RWK" and left(jobnumber,3)<>"TKB" then'normal job
	
	SQLSO="Select Quantity from tbl_MES_LotMasterSub where WipEntityName='"&jobnumber&"' and SheetNumber="&sheetnumber
	rsS.open SQLSO,connTicket,1,3
	
	if not rsS.eof then
		getAutoQuantity=rsS("Quantity")
	else
		getAutoQuantity=""
	end if
	rsS.close
	set rsS=nothing
else
	JOBNUMBER=jobnumber &"-"&string(3-len(sheetnumber),"0")&sheetnumber
	SQLS="select sum(QUANTITY) AS Quantity from PRINT_NEWJOB_HISTORY where NEW_JOBNUMBER='"&jobnumber&"'"
	'RESPONSE.WRITE sqls
	rsS.open SQLS,conn,1,3
	if not rsS.eof then
		getAutoQuantity=rsS("Quantity")
	ELSE
		getAutoQuantity=""
	end if
	rsS.close
	set rsS=nothing
end if

end function
%>