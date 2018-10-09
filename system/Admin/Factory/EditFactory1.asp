<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Factory/FactoryCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
factoryname=request.Form("factoryname")
SQL="select * from FACTORY where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("FACTORY_NAME")=factoryname
rs("TARGET_YIELD")=request.Form("finalyield")
rs("TARGET_FIRSTYIELD")=request.Form("firstyield")
rs("TARGET_LOST_QUANTITY")=request.Form("linelost_quantity")
rs("TARGET_LOST_AMOUNT")=request.Form("linelost_quantity")
rs("SCRAP_APPROVAL")=request.Form("scrap_approval")
rs("SCRAP_ADMINISTRATOR")=request.Form("scrap_administrator")
rs("FINANCE_TARGET_YIELD")=csng(request.Form("finance_target"))/100
rs("FINANCE_PLAN_TARGET_YIELD")=csng(request.Form("finance_plan_target"))/100
rs.update
rs.close
word="Successfully edit factory."
action="location.href='"&beforepath&"'"
else
word="Factory of "&sectionname&" has not existed, please input again."
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
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->