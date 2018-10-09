<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
code=request.QueryString("code")
SQL="select P1.CODE,P1.NAME,P1.ENGLISHNAME,PW.NTACCOUNT,PW.EMAIL from Personnel01 P1 inner join PersonnelWeb PW on P1.CODE=PW.CODE where P1.CODE='"&code&"' or ENGLISHNAME like '%"&code&"%'"
rskq.open SQL,conn_web,1,3
if not rskq.eof then
thiscode=rskq("CODE")
thisname=rskq("NAME")
english_name=rskq("ENGLISHNAME")
nt=rskq("NTACCOUNT")
email=rskq("EMAIL")
else
thiscode="None"
end if
rskq.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
parent.form1.code.value="<%=thiscode%>";
parent.form1.name.value="<%=english_name%>";
parent.form1.chinesename.value="<%=thisname%>";
parent.form1.NT.value="CORP\\<%=trim(nt)%>";
parent.form1.email.value="<%=trim(email)%>";
//parent.form1.factory.options[0].selected;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->