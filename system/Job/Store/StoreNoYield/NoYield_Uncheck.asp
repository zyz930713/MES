<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Store/StoreNoYieldCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="NoYield_Check.asp"
nid=trim(request.QueryString("nid"))
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
isupdate=false
'retest check in job_master_store
if NoYieldchecker=true then
	SQL="select * from JOB_MASTER_STORE where NID='"&nid&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	jobnumber=rs("JOB_NUMBER")
	part_number_tag=rs("PART_NUMBER_TAG")
	factory_id=rs("FACTORY_ID")
	line_name=rs("LINE_NAME")
	rs("NO_YIELD")=0
	isupdate=true
	rs.update
	word="复查成功！"
	else
	word="没有此数据！"
	end if
	rs.close
	'save a change history record in history table
	if isupdate=true then
	SQL="select * from JOB_MASTER_STORE_CHANGE"
	rs.open SQL,conn,1,3
	rs.addnew
	rs("STORE_NID")=nid
	rs("CHANGE_TYPE")="4"'1=change;2=delete;3=check no yield;4=uncheck no yield
	rs("CHANGE_CODE")=session("code")
	rs("CHANGE_TIME")=now()
	rs("CHANGE_REASON")="Cancel Yield Exclusion"
	rs("JOB_NUMBER")=jobnumber
	rs("PART_NUMBER_TAG")=part_number_tag
	rs("FACTORY_ID")=factory_id
	rs("LINE_NAME")=line_name
	rs.update
	rs.close
	end if
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