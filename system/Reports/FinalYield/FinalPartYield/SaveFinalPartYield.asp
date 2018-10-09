<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath="FinalPartYieldList.asp"
rnd_key=request.Form("rnd_key")
finalpart_name=trim(request.Form("finalpart_name"))
from_time=request.Form("from_time")
to_time=request.Form("to_time")
factory=request.Form("factory")
create_time=now()
factory_target_yield=request.Form("factory_target_yield")
NID="FP"&NID_SEQ("FINAL_PARTYIELD")
SQL="insert into FINAL_PARTYIELD_DETAIL select '"&NID&"',Part_NAME,INPUT_QUANTITY,OUTPUT_QUANTITY,FINAL_YIELD,STORE_NIDS,ASSEMBLY_INPUT_QUANTITY,ASSEMBLY_OUTPUT_QUANTITY,ASSEMBLY_YIELD from FINAL_PARTYIELD_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
SQL="select * from FINAL_PARTYIELD_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("REPORT_TIME")=create_time
rs("FINAL_PARTYIELD_NAME")=finalPart_name
rs("FACTORY_ID")=factory
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=from_time
rs("TO_TIME")=to_time
rs.update
rs.close
word="Successfully save FINAL Part YIELD report at "&create_time&"."
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