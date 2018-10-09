<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath="FinalSeriesYieldList.asp"
rnd_key=request.Form("rnd_key")
finalseries_name=trim(request.Form("finalseries_name"))
from_time=request.Form("from_time")
to_time=request.Form("to_time")
factory=request.Form("factory")
factory_target_yield=request.Form("factory_target_yield")
factory_target_firstyield=request.Form("factory_target_firstyield")
create_time=now()
NID="FS"&NID_SEQ("FINAL_SERIESYIELD")
SQL="insert into FINAL_SERIESYIELD_DETAIL select '"&NID&"',SERIES_NAME,INPUT_QUANTITY,OUTPUT_QUANTITY,FINAL_YIELD,STORE_NIDS,ASSEMBLY_INPUT_QUANTITY,ASSEMBLY_OUTPUT_QUANTITY,ASSEMBLY_YIELD,SERIES_TARGET_YIELD,OVERALL_EXCEPTION,SERIES_TARGET_FIRSTYIELD,SERIES_TARGET_INTERNALYIELD from FINAL_SERIESYIELD_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from FINAL_SERIESYIELD_DETAIL"
rs.open SQL,conn,1,3
rs.addnew
rs("FINAL_SERIESYIELD_ID")=NID
rs("SERIES_NAME")="OVERALL"
rs("INPUT_QUANTITY")=request.Form("total_assembly_input")
rs("OUTPUT_QUANTITY")=request.Form("total_final_output")
rs("FINAL_YIELD")=request.Form("total_final_yield")
rs("ASSEMBLY_INPUT_QUANTITY")=request.Form("total_assembly_input")
rs("ASSEMBLY_OUTPUT_QUANTITY")=request.Form("total_assembly_output")
rs("ASSEMBLY_YIELD")=request.Form("total_assembly_yield")
rs("SERIES_TARGET_YIELD")=factory_target_yield
rs("SERIES_TARGET_FIRSTYIELD")=factory_target_firstyield
rs.update
rs.close
SQL="select * from FINAL_SERIESYIELD_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("REPORT_TIME")=create_time
rs("FINAL_SERIESYIELD_NAME")=finalseries_name
rs("FACTORY_ID")=factory
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=from_time
rs("TO_TIME")=to_time
rs("FACTORY_TARGET_YIELD")=factory_target_yield
rs("FACTORY_TARGET_FIRSTYIELD")=factory_target_firstyield
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
word="Successfully save FINAL SERIES YIELD report at "&create_time&"."
action="location.href='"&beforepath&"'"
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
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