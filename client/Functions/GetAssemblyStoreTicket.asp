<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
Response.Charset = "GB2312"
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
</head>

<body>
<%
set rss=server.CreateObject("adodb.recordset")
ticketID=request.QueryString("ticketID")
SQL="select * from STORE_PRINT where NID='"&ticketID&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	if rs("Note")="Auto" then
		ids=rs("PRINT_MEMBERS")
		array_ids=split(ids,",")
		total=0
		j=0
		for i=0 to ubound(array_ids)
			SQLS="select * from JOB_MASTER_STORE_PRE where NID='"&array_ids(i)&"'"
			rsS.open SQLS,conn,1,3
			if not rsS.eof then
			total=total+cint(rsS("inspect_quantity"))
			j=j+1
			end if
			rsS.close
		next
		tasks=j&"条记录，总数："&total
	else
		tasks="确认页面选择错误！"
	end if
else
tasks="无"
end if
rs.close%>
<%=tasks%>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->