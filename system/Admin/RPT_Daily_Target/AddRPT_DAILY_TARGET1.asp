<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/RPT_DAILY_TARGET/RPT_DAILY_TARGETCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query



PRODUCT=trim(request("PRODUCT"))
LINE=trim(request("LINE"))
sqlN="select * from SUBSERIES where SUBSERIES_NAME='"&PRODUCT&"'"
rs.open sqlN,conn,1,3
if rs.eof then
word="子系列名称不存在！"
action="location.href='"&beforepath&"'"
end if
rs.close
sqlL="select * from LINE where LINE_NAME='"&LINE&"'"
rs.open sqlL,conn,1,3
if rs.eof then
word="线别名称不存在！"
action="location.href='"&beforepath&"'"
end if
rs.close

SQLT ="select *  from RPT_Daily_Target  where product='"&PRODUCT&"' and Line='"&Line&"'"

rs.open SQLT,conn,1,3
if not rs.eof then
word="此"&PRODUCT&"系列的"&Line&"线已存在！"
action="location.href='"&beforepath&"'"
end if
rs.close
if word="" then

TARGET=trim(request("TARGET"))
PROCESS_YIELD_TARGET=trim(request("PROCESS_YIELD_TARGET"))
ACOUSTIC_YIELD_TARGET=trim(request("ACOUSTIC_YIELD_TARGET"))
FOI_YIELD_TARGET=trim(request("FOI_YIELD_TARGET"))
LEAKAGE_YIELD_TARGET=trim(request("LEAKAGE_YIELD_TARGET"))
ACOUSITC_OFFSET=trim(request("ACOUSITC_OFFSET"))
FPY_TARGET=trim(request("FPY_TARGET"))
userCode=session("Code")
   
	
SQL="insert into RPT_DAILY_TARGET (PRODUCT,LINE,TARGET,PROCESS_YIELD_TARGET,ACOUSTIC_YIELD_TARGET,FOI_YIELD_TARGET,FPY_TARGET,LM_USER,LM_TIME,LEAKAGE_YIELD_TARGET) values('"&PRODUCT&"','"&LINE&"','"&TARGET&"','"&PROCESS_YIELD_TARGET&"','"&ACOUSTIC_YIELD_TARGET&"','"&FOI_YIELD_TARGET&"','"&FPY_TARGET&"','"&userCode&"',sysdate,'"&LEAKAGE_YIELD_TARGET&"')"
rs.open SQL,conn,1,3


word="Successfully save a New Daily Target."
action="location.href='"&beforepath&"'"
end if
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