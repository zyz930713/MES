<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Plan_Customer_Name_Config/Plan_Customer_Name_Config.aspCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query


CONFIGID=trim(request("CONFIGID"))
PART_NUMBER=trim(request("PART_NUMBER"))
CUSTOMER_NAME=trim(request("CUSTOMER_NAME"))
userCode=session("Code")
Action=trim(request("Action"))


if Action="Del" then
conn.Execute("Delete from Plan_Customer_Name_Config   where CONFIGID='"&CONFIGID&"'")
word="Successfully Del Plan_Customer_Name_Config."
action="history.back()"
else
SQL="select *  from Plan_Customer_Name_Config  where CONFIGID='"&CONFIGID&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	rs("PART_NUMBER")=PART_NUMBER
	rs("CUSTOMER_NAME")=CUSTOMER_NAME
	rs("LM_USER")=userCode
	rs("LM_TIME")=now()
rs.update
rs.close
word="Successfully edit Plan_Customer_Name_Config."
action="location.href='"&beforepath&"'"
else
word="Plan_Customer_Name_Config of "&PART_NUMBER&"  has not existed, please input again."
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