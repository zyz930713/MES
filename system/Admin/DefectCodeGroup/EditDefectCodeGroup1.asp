<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
group_name=request.Form("group_name")
defectcode_nids=replace(request.Form("toitem")," ","")
SQL="select * from DEFECTCODE_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("GROUP_NAME")=group_name
rs("GROUP_CHINESE_NAME")=request.Form("group_chinese_name")
rs("FACTORY_ID")=request.Form("factory")
rs("MEMBERS_ID")=defectcode_nids
rs.update
rs.close
word="Successfully edit DefectCode Group."
action="location.href='"&beforepath&"'"
else
word="DefectCode Group of "&group_name&" has not existed, please input again."
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