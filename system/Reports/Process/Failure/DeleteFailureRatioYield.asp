<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
factory_id=request("factory_id")
year_number=request("year_number")
week_number=request("week_number")
SQL="delete from BAR_REPORT.WEEKLY_FAILURE_RATIO where FACTORY_ID='"&factory_id&"' and YEAR_NUMBER='"&year_number&"' and WEEK_NUMBER='"&week_number&"'"
rs.open SQL,conn,1,3

word="Report is deleted successfully!"
action="location.href='"&beforepath&"'"
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