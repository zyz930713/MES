<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
finalfamily_name=request.QueryString("finalfamily_name")
SQL="select * from FINAL_FAMILYYIELD_LIST where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	byline=rs("BY_LINE")
	rs.delete
	rs.update
end if
rs.close
if byline="1" then
SQL="delete from FFAMILY_LINEYIELD_DETAIL where FINAL_FAMILYYIELD_ID='"&id&"'"
rs.open SQL,conn,3,3
else
SQL="delete from FINAL_FAMILYYIELD_DETAIL where FINAL_FAMILYYIELD_ID='"&id&"'"
rs.open SQL,conn,3,3
end if
word="Final Family Yield of "&finalfamily_name&" is deleted!"
%>
<script language="javascript">
alert("<%=word%>");
location.href='<%=beforepath%>';
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->