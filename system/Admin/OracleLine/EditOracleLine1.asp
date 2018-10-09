<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/OracleLine/OracleLineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
line_id=request.Form("line_id")
SQL="select * from tbl_Wip_Line_Sub where LINE_ID='"&line_id&"'"
rsPR.open SQL,connPR,1,3
if rsPR.eof then
rsPR.addnew
end if
rsPR("Line_Id")=line_id
rsPR("SubQuantity")=request.Form("subqty")
rsPR.update
rsPR.close
word="Successfully edit Line."
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
<!--#include virtual="/WOCF/POCF_Close.asp" -->