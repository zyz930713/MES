<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
defectcodename=request.QueryString("defectcodename")
SQL="update DEFECTCODE_NEW set IS_DELETE=1,LASTUPDATE_PERSON='"&session("code")&"',LASTUPDATE_TIME='"&date()&"' where NID='"&id&"'"
rs.open SQL,conn,3,3
SystemLog "Admin - DefectCode","Delete DefectCode of "&defectcodename&" ("&id&")"
word="Defect Code of "&defectcodename&" is deleted!"
action="location.href='"&beforepath&"'"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->