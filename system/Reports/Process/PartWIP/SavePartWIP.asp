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
beforepath="PartWIPList.asp"
rnd_key=request.Form("rnd_key")
part_WIP_name=request.Form("part_WIP_name")
create_time=now()
NID="PW"&NID_SEQ("PART_WIP")
SQL="insert into PART_WIP_DETAIL select '"&NID&"',PART_NUMBER_TAG,STATION_ID,QUANTITY,JOB_NUMBERS1,JOB_NUMBERS2,JOB_NUMBERS3,JOB_NUMBERS4,JOB_NUMBERS5,JOB_NUMBERS6,JOB_NUMBERS7,JOB_NUMBERS8,JOB_NUMBERS9,JOB_NUMBERS10 from PART_WIP_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATE_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from PART_WIP_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("SECTION_ID")=request.Form("section")
rs("REPORT_TIME")=create_time
rs("PART_WIP_NAME")=part_WIP_name
rs("CREATOR_CODE")=session("code")
rs.update
rs.close
word="Successfully save Part WIP report at "&create_time&"."
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