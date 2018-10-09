<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
NID=request.QueryString("NID")
ids=""
SQL="select * from STORE_PRINT where NID='"&NID&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("PRINT_TIME")=now()
ids=rs("PRINT_MEMBERS")
rs.update
end if
rs.close
' FACTORY_ID='FA00000002' and 
SQL="select * from JOB_MASTER_STORE_Pre where NID in ('"&replace(ids,",","','")&"') order by PART_NUMBER_TAG"
rs.open SQL,conn,1,3
if not rs.eof then
	idcount=rs.recordcount
	while not rs.eof
		rs("PRINT_TIMES")=1
		'rs.update
		rs.movenext
	wend
end if
rs.close
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>产品入库清单更新</title>
<link href="/CSS/Print.css" rel="stylesheet" type="text/css">
<script language="javascript">
/*parent.document.all.OK.disabled=true;
parent.document.all.Print.disabled=false;
parent.document.all.Close.disabled=false;*/
parent.location.href="PrintStoreRepeatList.asp?printid=<%=NID%>";
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->