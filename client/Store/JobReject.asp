<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="JobReject.asp"
jobnumber=request.QueryString("jobnumber")
reason=request.QueryString("reason")

SQL="select LINE_NAME from JOB where ROWNUM=1 and JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
line=rs("LINE_NAME")
	set rsLine=server.CreateObject("adodb.recordset")
	SQLLine="select EMAIL from USERS where USER_CODE=(select LEADER from LINE where LINE_NAME='"&line&"')"
	rsLine.open SQLLine,conn,1,3
	if not rsLine.eof then
	email=rsLine("EMAIL")
	end if
	rsLine.close
	set rsLine=nothing
	if email<>"" then
		sendEmail "Barcode.System@knowles.com",trim(email),"<"&jobnumber&">��ⷢ���˻أ�ԭ����"&reason&"��",line&"��ࣺ<br><"&jobnumber&">��ⷢ���˻أ�ԭ����"&reason&"��<br>Barcode System"
		word="���͵����ʼ�֪ͨ"&email&"���Զ��˻سɹ���"
	else
	word="û��������Ϣ���Զ��˻�ʧ�ܣ�"
	end if
else
word="û�иù��������Զ��˻�ʧ�ܣ�"
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>")
window.close()
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->