<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
NID=request.QueryString("NID")
ids=""
SQL="SELECT PRINT_TIME,PRINT_MEMBERS FROM STORE_PRINT WHERE NID='"&NID&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	rs("PRINT_TIME")=now()
	ids=rs("PRINT_MEMBERS")
	rs.update
end if
rs.close

SQL="UPDATE JOB_PACK_DETAIL SET IS_PRINTED='1' WHERE BOX_ID IN ('"&replace(ids,",","','")&"') "
rs.open SQL,conn,1,3
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript">
	window.close();
</script>
</head>

<body>
</body>
</html>

