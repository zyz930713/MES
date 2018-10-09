<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Packing_Plan/Packing_PlanCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query



PLAN_DATE=trim(request("PLAN_DATE"))
PART_NUMBER=trim(request("PART_NUMBER"))
QUANTITY=trim(request("QUANTITY"))
CUSTOMER_NAME=trim(request("CUSTOMER_NAME"))
CUSTOMER_PART_NUMBER=trim(request("CUSTOMER_PART_NUMBER"))
Remark=trim(request("Remark"))
PRIORITY=trim(request("PRIORITY"))
DELIVERY_TIME=trim(request("DELIVERY_TIME"))
userCode=session("Code")
set rs=server.createobject("adodb.recordset")
	
SQL="insert into PACKING_PLAN (PLAN_ID,PLAN_DATE,PART_NUMBER,QUANTITY,CUSTOMER_NAME,CUSTOMER_PART_NUMBER,REMARK,PRIORITY,DELIVERY_TIME,COMPLETE_TIME,STATUS,LM_USER,LM_TIME) values('PL'||to_char(sysdate,'ymmddhh24miss'),'"&PLAN_DATE&"','"&PART_NUMBER&"','"&QUANTITY&"','"&CUSTOMER_NAME&"','"&CUSTOMER_PART_NUMBER&"','"&Remark&"','"&PRIORITY&"','"&DELIVERY_TIME&"',null,'Initial','"&userCode&"',sysdate)"



rs.open SQL,conn,1,3


word="Successfully save a New Packing Plan."
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