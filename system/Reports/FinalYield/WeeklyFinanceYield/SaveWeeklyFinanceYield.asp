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
beforepath="FinalFamilyYieldList.asp"
rnd_key=request.Form("rnd_key")
finance_yield_name=trim(request.Form("finance_yield_name"))
from_time=request.Form("from_time")
to_time=request.Form("to_time")
factory=request.Form("factory")
factory_target_yield=request.Form("factory_target_yield")
factory_plan_target_yield=request.Form("factory_plan_target_yield")
create_time=now()
NID="FF"&NID_SEQ("FINAL_FAMILYYIELD")
SQL="insert into WEEKLY_FINANCE_YIELD_DETAIL select '"&NID&"',FAMILY_NAME,INPUT_QUANTITY,OUTPUT_QUANTITY,SCRAP_QUANTITYI,NPUT_AMOUNT,OUTPUT_AMOUNT,SCRAP_AMOUNT,AMOUNT_YIELD,TARGET,PLAN_TARGET from WEEKLY_FINANCE_YIELD_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from WEEKLY_FINANCE_YIELD_DETAIL"
rs.open SQL,conn,1,3
rs.addnew
rs("FINANCE_YIELD_ID")=NID
rs("FAMILY_NAME")="OVERALL"
rs("INPUT_AMOUNT")=request.Form("total_input_amount")
rs("OUTPUT_AMOUNT")=request.Form("total_output_amount")
rs("SCRAP_AMOUNT")=request.Form("total_scrap_amount")
rs("FINAL_YIELD")=request.Form("total_yield")
rs("TARGET_YIELD")=factory_target_yield
rs("PLAN_TARGET_YIELD")=factory_plan_target_yield
rs.update
rs.close
SQL="select * from WEEKLY_FINANCE_YIELD_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("REPORT_TIME")=create_time
rs("WEEKLY_FINANCEYIELD_NAME")=finance_yield_name
rs("FACTORY_ID")=factory
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=from_time
rs("TO_TIME")=to_time
rs("FACTORY_TARGET_YIELD")=factory_target_yield
rs("FACTORY_PLAN_TARGET_YIELD")=factory_plan_target_yield
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
word="Successfully save WEEKLY FINANCE YIELD report at "&create_time&"."
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