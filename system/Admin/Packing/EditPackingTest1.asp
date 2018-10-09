<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
PART_ID=trim(request("PART_ID"))
PART_NUMBER=trim(request("PART_NUMBER"))
TEST_NAME=trim(request("TEST_NAME"))
ALARM_MSG=trim(request("ALARM_MSG"))
Test_Type=trim(request("Test_Type"))
Action=trim(request("Action"))


if Action="Del" then



conn.Execute("Delete from part_test_name  where  PART_ID='"&PART_ID&"'")
word="Successfully Del Packing Config."
action="history.back()"


else
conn.Execute("update part_test_name  set TEST_NAME='"&TEST_NAME&"',alarm_msg='"&ALARM_MSG&"',Test_Type='"&Test_Type&"' where  PART_ID='"&PART_ID&"'")
word="Successfully edit Packing Config."
action="location.href='"&beforepath&"'"

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