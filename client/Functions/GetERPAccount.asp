<%
function getERPAccount()
set rsS=server.CreateObject("adodb.recordset")
SQLD="Select * from REMOTE_GL_ACCOUNT order by ACCOUNT_ALIAS"
rsS.open SQLD,conn,1,3
while not rsS.eof 
	optiontext=optiontext&"<option value='"&rsS("GCCID")&"'>"&rsS("ACCOUNT_ALIAS")&"</option>"
rsS.movenext
wend
getERPAccount=optiontext
rsS.close
set rsS=nothing
end function
 
function getERPAccountName(GCCID)
set rsS=server.CreateObject("adodb.recordset")
SQLD="Select * from REMOTE_GL_ACCOUNT  where GCCID='"+GCCID+"'"
rsS.open SQLD,conn,1,3
if(rsS.recordcount>0) then
getERPAccountName=rsS("ACCOUNT_ALIAS")
else
getERPAccountName=""
end if 
rsS.close
set rsS=nothing
end function
%>