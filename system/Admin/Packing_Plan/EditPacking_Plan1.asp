<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Packing_Plan/Packing_PlanCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query


PLAN_ID=request.Form("PLAN_ID")
PLAN_DATE=trim(request("PLAN_DATE"))
PART_NUMBER=trim(request("PART_NUMBER"))
QUANTITY=trim(request("QUANTITY"))
CUSTOMER_NAME=trim(request("CUSTOMER_NAME"))
CUSTOMER_PART_NUMBER=trim(request("CUSTOMER_PART_NUMBER"))
REMARK=trim(request("REMARK"))
PRIORITY=trim(request("PRIORITY"))
DELIVERY_TIME=trim(request("DELIVERY_TIME"))
NewStatus=trim(request("NewStatus"))
if NewStatus="Complete" then
COMPLETE_TIME=now()
end if
userCode=session("Code")

SQL="select * from PACKING_PLAN where PLAN_ID='"&PLAN_ID&"'"

rs.open SQL,conn,1,3
if not rs.eof then	
	rs("PLAN_DATE")=PLAN_DATE
	rs("PART_NUMBER")=PART_NUMBER
	rs("QUANTITY")=QUANTITY
	rs("CUSTOMER_NAME")=CUSTOMER_NAME
	rs("CUSTOMER_PART_NUMBER")=CUSTOMER_PART_NUMBER
	rs("REMARK")=REMARK
	rs("PRIORITY")=PRIORITY
	rs("DELIVERY_TIME")=DELIVERY_TIME	
	rs("Status")=NewStatus
	rs("COMPLETE_TIME")=COMPLETE_TIME
	rs("LM_USER")=userCode
	rs("LM_TIME")=now()
rs.update
rs.close

'
'if NewStatus="Cancel" then
'sql= "update job_pack_detail set pallet_id='',customer_label_id='',plan_id='' where plan_id='"&PLAN_ID&"'"
'
'conn.execute sql
'end if


word="Successfully edit Packing Plan."
action="location.href='"&beforepath&"'"
else
word="Packing Plan of "&PART_NUMBER&" has not existed, please input again."
action="history.back()"
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