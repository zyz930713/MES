<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
start_quantity=0
SQLPR="select Quantity from tbl_MES_LotMasterSub where WipEntityName='"&jobnumber&"' and SheetNumber='"&sheetnumber&"'"
rsPR.open SQLPR,connPR,1,3
if not rsPR.eof then
start_quantity=cint(rsPR("Quantity"))
end if
rsPR.close

SQL="select JOB_START_QUANTITY from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	if isnull(rs("JOB_START_QUANTITY")) or rs("JOB_START_QUANTITY")="" then
		rs("JOB_START_QUANTITY")=start_quantity
		word="Successful to update new quantity!"
		rs.update
	else
		if cint(rs("JOB_START_QUANTITY"))<>start_quantity and start_quantity<>0 then
		rs("JOB_START_QUANTITY")=start_quantity
		word="Successful to update new quantity!"
		rs.update
		else
		word="Do not need to udpate!"
		end if
	end if
end if
rs.close
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->