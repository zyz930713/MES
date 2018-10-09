<%
function getJobSetStartQuantity(jobnumber)	
set rsJ=server.CreateObject("adodb.recordset")
SQLJ="Select START_QUANTITY from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
rsJ.open SQLJ,conn,1,3
if not rsJ.eof then
getJobSetStartQuantity=cLnG(rsJ("START_QUANTITY"))
else
getJobSetStartQuantity=0
end if
rsJ.close
set rsJ=nothing
end function
%>