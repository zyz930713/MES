<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
station=request.QueryString("station")
SQL="select NID,ACTION_NAME from ACTION where STATION_ID='"&station&"' or STATION_ID is null order by ACTION_NAME"
rs.open SQL,conn,1,3
if not rs.eof then
while not rs.eof 
options=options&rs("NID")&"*"&rs("ACTION_NAME")&"|"
rs.movenext
wend
end if
rs.close
options=left(options,len(options)-1)
%>
<%=options%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->