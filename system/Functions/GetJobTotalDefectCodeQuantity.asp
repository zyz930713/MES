<%
function getJobTotalDefectCodeQuantity(jobnumber,sheetnumber)	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
	getJobTotalDefectCodeQuantity=getJobTotalDefectCodeQuantity+cint(rsS("DEFECT_QUANTITY"))
	rsS.movenext
	wend
else
getJobTotalDefectCodeQuantity=0
end if
rsS.close
set rsS=nothing
end function
%>