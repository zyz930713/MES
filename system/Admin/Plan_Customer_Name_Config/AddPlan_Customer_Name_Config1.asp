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




PART_NUMBER=trim(request("PART_NUMBER"))

CUSTOMER_NAME=trim(request("CUSTOMER_NAME"))
Set TypeLib = CreateObject("Scriptlet.TypeLib")
Guid = TypeLib.Guid
guid= Mid(Guid,2,36)
response.Write(guid)
'response.End()

userCode=session("Code")
set rs=server.createobject("adodb.recordset")
	
SQL="insert into Plan_Customer_Name_Config (CONFIGID,PART_NUMBER,CUSTOMER_NAME,LM_USER,LM_TIME) values('"&Guid&"','"&PART_NUMBER&"','"&CUSTOMER_NAME&"','"&userCode&"',sysdate)"



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