<%
function getERPAccount()
set rsS=server.CreateObject("adodb.recordset")
SQLD="Select * from REMOTE_GL_ACCOUNT where ACCOUNT_ALIAS IN ('KE.SCRAP', 'ASSY.SCRAP', 'VALADD.SCRAP', 'DELTEK.SCRAP') order by ACCOUNT_ALIAS"
rsS.open SQLD,conn,1,3
while not rsS.eof 
	optiontext=optiontext&"<option value='"&rsS("GCCID")&"'>"&rsS("ACCOUNT_ALIAS")&"</option>"
rsS.movenext
wend
getERPAccount=optiontext
rsS.close
set rsS=nothing
end function
%>