<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobTracking/JobTrackingCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
base_name=request.QueryString("base_name")
excelfilename=request.QueryString("excelfilename")
SQL="delete from JOB_BASE_LIST where NID='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from JOB_BASE_DETAIL where BASE_ID='"&id&"'"
rs.open SQL,conn,3,3
	set myFs=server.createObject("scripting.FileSystemObject") 
	ICfilePath=server.mappath("\Reports\Tracking\BaseTracking\EXCELs")
	myfile=ICfilePath&"\"&filename 
	if myFs.FileExists(myfile)=true then
	myFs.DeleteFile myfile
	end if
	set myFs=nothing
word="Base Report of "&base_name&" is deleted!"
%>
<script language="javascript">
alert("<%=word%>");
location.href='<%=beforepath%>';
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->