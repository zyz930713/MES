<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Material/MaterialCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
materialnumber=request.Form("materialnumber")
SQL="select * from MATERIAL where MATERIAL_NUMBER='"&materialnumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="MT"&NID_SEQ("MATERIAL")
rs("MATERIAL_NUMBER")=materialnumber
rs("MATERIAL_NAME")=request.Form("materialname")
rs("FACTORY_ID")=request.Form("factory")
rs("UNIT")=request.Form("materialunit")
rs("MIN_UNIT_RATIO")=request.Form("ratio")
if trim(request.Form("locked"))="" then
rs("LOCKED_LOT")=null
else
rs("LOCKED_LOT")=replace(trim(request.Form("locked"))," ","")
end if
rs.update
word="Successfully save a New Material."
action="location.href='"&beforepath&"'"
else
word="Material of "&materialnumber&" has existed, please input again."
action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->