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
beforepath="OutputList.asp"
rnd_key=request.Form("rnd_key")
output_name=request.Form("output_name")
create_time=now()
NID="OU"&NID_SEQ("OUTPUT")
SQL="insert into OUTPUT_DETAIL select '"&NID&"',LINE_ID,STATION_ID,QUANTITY,JOB_NUMBERS from OUTPUT_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATE_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from OUTPUT_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("SECTION_ID")=request.Form("section")
rs("REPORT_TIME")=create_time
rs("OUTPUT_NAME")=output_name
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=request.Form("fromtime")
rs("TO_TIME")=request.Form("totime")
rs.update
rs.close
word="Successfully save Output report at "&create_time&"."
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