<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
PART_ID=request("PART_ID")
TEST_NAME=request("TEST_NAME")
PART_NUMBER=request("PART_NUMBER")
path=request.QueryString("path")
query=request.QueryString("query")

set rs1=server.CreateObject("adodb.recordset")
SQL="select * from part_test_name  where PART_ID='"&PART_ID&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/PACKING/EditPackingTest1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td height="20"><span id="inner_PartNumber"></span> <span class="red">*</span></td>
  <td height="20"> <input name="PART_NUMBER" type="text" id="PART_NUMBER"  value="<%=rs("PART_NUMBER")%>" size="15" readonly></td>
</tr>
<tr>
  <td height="20"><span id="td_TestName"></span> <span class="red">*</span></td>
  <td height="20"><input name="TEST_NAME" type="text" id="TEST_NAME"  value="<%=rs("TEST_NAME")%>" size="10" ></td>
</tr>
<tr>
  <td height="20"><span id="td_AlarmMsg"></span> <span class="red">*</span></td>
  <td height="20"><input name="ALARM_MSG" type="text" id="ALARM_MSG"  value="<%=rs("ALARM_MSG")%>" size="40"></td>
</tr>
<tr>
  <td height="20"><span id="td_testType"></span> <span class="red">*</span></td>
  <td height="20"><input name="Test_Type" type="text" id="Test_Type"  value="<%=rs("Test_Type")%>" size="40"></td>
</tr>




  <tr>
    <td height="20" colspan="2"><div align="center">
    <input name="PART_ID" type="hidden" id="PART_ID" value="<%=PART_ID%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input name="Reset" type="reset" id="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->