<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Material/MaterialCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
materialnumber=request.Form("materialnumber")
SQL="select * from MATERIAL where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
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
word="Successfully edit Material."
action="location.href='"&beforepath&"'"
else
word="Material of "&materialnumber&" has not existed, please input again."
action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->