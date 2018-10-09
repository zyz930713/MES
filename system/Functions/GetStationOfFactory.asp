<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.QueryString("factory")
SQL="select NID,STATION_NAME from STATION where FACTORY_ID='"&factory&"' order by STATION_NAME"
rs.open SQL,conn,1,3
if not rs.eof then
while not rs.eof
'station_name=replace(rs("STATION_NAME"),"/","//")
station_name=replace(rs("STATION_NAME"),"-"," ")
options=options&rs("NID")&"*"&station_name&"|"
rs.movenext
wend
end if
rs.close
options=left(options,len(options)-1)
%>
<%=options%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->