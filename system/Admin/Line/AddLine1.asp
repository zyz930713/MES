<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
linename=trim(request.Form("linename"))
SQL="select * from LINE where LINE_NAME='"&linename&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="LI"&NID_SEQ("LINE")
rs("LINE_NAME")=linename
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
rs("MACHINE_LABELS")=REQUEST.Form("labels")
rs("FACTORY_CODE")=request("FACTORY_CODE")
rs("CODE_LINENAME")=trim(request("CODE_LINENAME"))
rs("CODE_NAME")=trim(request("CODE_NAME"))
rs("CODE_NAME2")=trim(request("CODE_NAME2"))
rs("Code_Date690")=trim(request("Code_Date690"))
rs("Code_Date")=request("Code_Date")
rs("VERSION_NUMBER")=request("VERSION_NUMBER")
rs("PRODUCT")=request("PRODUCT")
rs.update
rs.close
word="Successfully save a New Line."
action="location.href='"&beforepath&"'"
else
word="Line of "&linename&" has existed, please input again."
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