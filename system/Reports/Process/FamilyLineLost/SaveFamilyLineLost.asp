<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath="FamilyLineLostList.asp"
rnd_key=request.Form("rnd_key")
finalfamily_name=trim(request.Form("finalfamily_name"))
from_time=request.Form("from_time")
to_time=request.Form("to_time")
factory=request.Form("factory")
factory_target_quantity=request.Form("factory_target_quantity")
factory_target_amount=request.Form("factory_target_amount")
create_time=now()
NID="FL"&NID_SEQ("FAMILY_LINELOST")
SQL="insert into FAMILY_LINELOST_DETAIL select '"&NID&"',FAMILY_NAME,INPUT_QUANTITY,LINELOST_QUANTITY,LINELOST_PERCENTAGE,INCLUDED_SYSTEM_ITEMS,LINELOST_AMOUNT,FAMILY_TARGET_QUANTITY,FAMILY_TARGET_AMOUNT from FAMILY_LINELOST_DETAIL_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
'SQL="select * from FAMILY_LINELOST_DETAIL"
'rs.open SQL,conn,1,3
'rs.addnew
'rs("FAMILY_LINELOST_ID")=NID
'rs("FAMILY_NAME")="OVERALL"
'rs("INPUT_QUANTITY")=csng(request.Form("total_input"))
'rs("LINELOST_QUANTITY")=csng(request.Form("total_lost_quantity"))
'rs("LINELOST_AMOUNT")=csng(request.Form("total_lost_amount"))
'rs("LINELOST_PERCENTAGE")=0
'rs("FAMILY_TARGET_QUANTITY")=factory_target_quantity
'rs("FAMILY_TARGET_AMOUNT")=factory_target_amount
'rs.update
'rs.close
SQL="select * from FAMILY_LINELOST_LIST"
rs.open SQL,conn,1,3
rs.addnew
rs("NID")=NID
rs("REPORT_TIME")=create_time
rs("FAMILY_LINELOST_NAME")=finalfamily_name
rs("FACTORY_ID")=factory
rs("CREATOR_CODE")=session("code")
rs("FROM_TIME")=from_time
rs("TO_TIME")=to_time
rs("FACTORY_TARGET_QUANTITY")=factory_target_quantity
rs("FACTORY_TARGET_AMOUNT")=factory_target_amount
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
word="Successfully save FAMILY LINE LOST report at "&create_time&"."
action="location.href='"&beforepath&"'"
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->