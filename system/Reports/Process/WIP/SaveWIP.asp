<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath="WIPList.asp"
rnd_key=request.Form("rnd_key")
WIP_name=request.Form("WIP_name")
create_time=now()
NID="WI"&NID_SEQ("WIP")
SQL="insert into WIP_DETAIL select '"&NID&"',LINE_ID,STATION_ID,QUANTITY,JOB_NUMBERS1,JOB_NUMBERS2,JOB_NUMBERS3,JOB_NUMBERS4,JOB_NUMBERS5,JOB_NUMBERS6,JOB_NUMBERS7,JOB_NUMBERS8,JOB_NUMBERS9,JOB_NUMBERS10 from WIP_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATE_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from WIP_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("SECTION_ID")=request.Form("section")
rs("REPORT_TIME")=create_time
rs("WIP_NAME")=WIP_name
rs("CREATOR_CODE")=session("code")
rs.update
rs.close
word="Successfully save WIP report at "&create_time&"."
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->