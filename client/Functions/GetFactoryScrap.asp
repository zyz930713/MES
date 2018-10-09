<%
function getFactoryScrap(factory)
set rsF=server.CreateObject("adodb.recordset")
SQLF="Select SCRAP_APPROVAL from FACTORY where NID='"&factory&"'"
rsF.open SQLF,conn,1,3
if not rsF.eof then
	if isnull(rsF("SCRAP_APPROVAL"))=false and rsF("SCRAP_APPROVAL")<>"" then
	getFactoryScrap=cint(rsF("SCRAP_APPROVAL"))/100
	else
	getFactoryScrap=0
	end if
else
getFactoryScrap=0
end if
rsF.close
set rsF=nothing
end function
%>