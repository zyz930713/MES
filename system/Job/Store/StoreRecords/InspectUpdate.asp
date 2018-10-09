<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsStockAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="InspectUpdate.asp"
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
nid=trim(request.QueryString("nid"))
new_inspect_quantity=cint(request.QueryString("new_inspect"))
if admin=true then
	SQL="select JOB_NUMBER,INSPECT_QUANTITY from JOB_MASTER_STORE where NID='"&nid&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	job_number=rs("JOB_NUMBER")
	old_inspect_quantity=csng(rs("INSPECT_QUANTITY"))
	rs("INSPECT_QUANTITY")=new_inspect_quantity
		if old_inspect_quantity<>new_inspect_quantity then
		rs.update
		root_update=true
		else
		root_update=false
		end if
	end if
	rs.close
	if root_update=true then
		SQL="select FINAL_INSPECT_QUANTITY,FINAL_SCRAP_QUANTITY,CONFIRM_GOOD_QUANTITY from JOB_MASTER where JOB_NUMBER='"&job_number&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
		final_inspect_quantity=rs("FINAL_INSPECT_QUANTITY")
		rs("FINAL_INSPECT_QUANTITY")=final_inspect_quantity+(new_inspect_quantity-old_inspect_quantity)
		rs("FINAL_SCRAP_QUANTITY")=final_scrap_quantity+(old_inspect_quantity-new_inspect_quantity)
		rs("CONFIRM_GOOD_QUANTITY")=final_inspect_quantity+(new_inspect_quantity-old_inspect_quantity)
		rs.update
		end if
		rs.close
	end if
word="更新成功！"
else
word="无权操作！"
end if
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
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