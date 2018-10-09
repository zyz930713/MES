<%
Function GetUserAgent(UserCode)
	set objRst=server.CreateObject("adodb.recordset")
	SQL="select * from PERSONNELWEB where CODE='" & UserCode&"'"
	objRst.open SQL,conn_web,1,3
	if not objRst.eof then
		GetUserAgent=objRst("Agent")
	end if
	set objRst=nothing
End Function
%>