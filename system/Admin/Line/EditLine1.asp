<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
linename=request.Form("linename")
SQL="select * from LINE where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("LINE_NAME")=linename
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
rs("LEADER")=request.Form("leader")
rs("SUPERVISOR")=request.Form("supervisor")
rs("MACHINE_LABELS")=REQUEST.Form("labels")
rs("FACTORY_CODE")=request("FACTORY_CODE")
rs("CODE_LINENAME")=trim(request("CODE_LINENAME"))
rs("CODE_NAME")=trim(request("CODE_NAME"))
rs("CODE_NAME2")=trim(request("CODE_NAME2"))
rs("CODE_Date690")=trim(request("code_Date690"))
rs("CODE_Date")=trim(request("code_Date"))
rs("VERSION_NUMBER")=request("VERSION_NUMBER")
rs("PRODUCT")=request("PRODUCT")
rs.update
rs.close
word="Successfully edit Line."
action="location.href='"&beforepath&"'"
else
word="Line of "&linename&" has not existed, please input again."
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