<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/FinanceSeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
seriesgroupname=request.Form("seriesgroupname")
SQL="select * from FINANCE_SERIES_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("SERIES_GROUP_NAME")=seriesgroupname
rs("FACTORY_ID")=request.Form("factory")
rs("PREFIX")=trim(request.Form("prefix"))
rs("TARGET_YIELD")=request.Form("yield")/100
rs("PLAN_TARGET_YIELD")=request.Form("plan_yield")/100
rs.update
rs.close
SQL="update FINANCE_SERIES set SERIES_GROUP_ID='"&id&"' where instr('"&request.Form("toitem")&"',NID)>0"
rs.open SQL,conn,1,3
SQL="update PRODUCT_MODEL set SERIES_GROUP_ID='"&id&"' where instr('"&request.Form("toitem")&"',SERIES_ID)>0"
rs.open SQL,conn,1,3
word="Successfully edit Series Group."
action="location.href='"&beforepath&"'"
else
word="Series Group of "&seriesgroupname&" has not existed, please input again."
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