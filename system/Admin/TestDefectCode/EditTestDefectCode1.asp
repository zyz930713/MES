<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TestDefectCode/TestDefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
defectname=trim(request.Form("defectname"))
SQL="select * from TEST_DEFECTCODE where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("DEFECT_NAME")=defectname
rs("VALUE_TYPE")=request.Form("value_type")
rs("SCALE")=trim(request.Form("scale"))
rs.update
word="Successfully edit Test Defect Code."
action="location.href='"&beforepath&"'"
else
word="Test Defect Code of "&defectname&" has existed, please input again."
action="history.back()"
end if
rs.close
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