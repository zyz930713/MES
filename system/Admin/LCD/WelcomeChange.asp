<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
thisday=replace(request.QueryString("thisday"),"_","-")
SQL="select * from DAILY_WELCOME where to_char(WELCOME_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("WELCOME_DAY")=thisday
end if
rs("WELCOME_TYPE")=request.QueryString("thistype")
rs("CREATOR_CODE")=session("code")
rs.update
rs.close
if request.QueryString("thistype")="1" then
word="Successfully change settings."
action="history.back()"
else
word="Successfully change settings. Please upload file."
action="location.href('Welcome.asp?thisday="&thisday&"&path="&path&"&query="&query&"')"
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