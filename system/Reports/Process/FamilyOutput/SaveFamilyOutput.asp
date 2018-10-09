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
beforepath="FamilyOutputList.asp"
rnd_key=request.Form("rnd_key")
family_output_name=request.Form("family_output_name")
create_time=now()
NID="FO"&NID_SEQ("FAMILY_OUTPUT")
SQL="insert into FAMILY_OUTPUT_DETAIL select '"&NID&"',FAMILY_NAME,STATION_NAME,OUTPUT from FAMILY_OUTPUT_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from FAMILY_OUTPUT_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
rs("REPORT_TIME")=create_time
rs("FAMILY_OUTPUT_NAME")=family_output_name
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=request.Form("fromtime")
rs("TO_TIME")=request.Form("totime")
if request.Form("isww")="1" then
rs("CHART_WEEK")=request.Form("wwnumber")
rs("CHART_YEAR")=request.Form("yearnumber")
rs("CHART_MONTH")=request.Form("monthnumber")
else
rs("CHART_WEEK")=null
rs("CHART_YEAR")=null
rs("CHART_MONTH")=null
end if
rs.update
rs.close
word="Successfully save Family Output report at "&create_time&"."
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