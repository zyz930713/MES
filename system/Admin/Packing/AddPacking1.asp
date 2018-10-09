<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
PART_NUMBER=trim(request("PART_NUMBER"))
TEST_NAME=trim(request("TEST_NAME"))
ALARM_MSG=trim(request("ALARM_MSG"))
Test_Type=trim(request("Test_Type"))
SQL="select * from part_test_name where test_name='"&TEST_NAME&"' and  PART_NUMBER='"&PART_NUMBER&"' and Test_Type='"&Test_Type&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("Part_ID")="P"&NID_SEQ("Part_New_SEQ")
rs("TEST_NAME")=TEST_NAME
rs("PART_NUMBER")=PART_NUMBER
rs("ALARM_MSG")=ALARM_MSG
rs("Test_Type")=Test_Type
rs.update
rs.close
word="Successfully save a New Packing config."
action="location.href='"&beforepath&"'"
else
word="Packing config of "&TEST_NAME&" has existed, please input again."
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