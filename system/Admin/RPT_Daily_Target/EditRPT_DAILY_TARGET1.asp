<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/RPT_Daily_Target/RPT_DAILY_TARGETCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query


PRODUCT=trim(request("PRODUCT"))
LINE=trim(request("LINE"))
TARGET=trim(request("TARGET"))
PROCESS_YIELD_TARGET=trim(request("PROCESS_YIELD_TARGET"))
ACOUSTIC_YIELD_TARGET=trim(request("ACOUSTIC_YIELD_TARGET"))
FOI_YIELD_TARGET=trim(request("FOI_YIELD_TARGET"))
LEAKAGE_YIELD_TARGET=trim(request("LEAKAGE_YIELD_TARGET"))
ACOUSITC_OFFSET=trim(request("ACOUSITC_OFFSET"))
FPY_TARGET=trim(request("FPY_TARGET"))
userCode=session("Code")
Action=trim(request("Action"))


if Action="Del" then

conn.Execute("Delete from RPT_Daily_Target   where product='"&PRODUCT&"' and Line='"&Line&"'")
word="Successfully Del RPT_Daily_Target."
action="history.back()"

else
SQL="select *  from RPT_Daily_Target  where product='"&PRODUCT&"' and Line='"&Line&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	
	
	
	
	rs("PRODUCT")=PRODUCT
	rs("LINE")=LINE
	rs("TARGET")=TARGET
	rs("PROCESS_YIELD_TARGET")=PROCESS_YIELD_TARGET
	rs("ACOUSTIC_YIELD_TARGET")=ACOUSTIC_YIELD_TARGET
	rs("FOI_YIELD_TARGET")=FOI_YIELD_TARGET
	rs("LEAKAGE_YIELD_TARGET")=LEAKAGE_YIELD_TARGET
	rs("ACOUSITC_OFFSET")=ACOUSITC_OFFSET
	rs("FPY_TARGET")=FPY_TARGET	
	rs("LM_USER")=userCode
	rs("LM_TIME")=now()
rs.update
rs.close






word="Successfully edit RPT_Daily_Target."
action="location.href='"&beforepath&"'"
else
word="RPT_Daily_Target of "&PRODUCT&"  has not existed, please input again."
action="history.back()"
end if
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