<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Model/ModelCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select P.*,F.FACTORY_NAME from PRODUCT_MODEL P inner join FACTORY F on P.FACTORY_ID=F.NID where P.ITEM_ID='"&id&"'"
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
<form action="/Admin/Model/EditModel1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
<tr>
  <td width="153" height="20"><span id="inner_PartNumber"></span></td>
    <td width="601" height="20" class="red">
      <div align="left"><%=rs("ITEM_NAME")%></div></td>
    </tr>
<tr>
  <td height="20"><span id="td_Description"></span></td>
  <td height="20"><%=rs("DESCRIPTION")%></td>
</tr>
<tr>
  <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
  <td height="20"><%=rs("FACTORY_NAME")%></td>
</tr>
<tr>
  <td height="20">CCL*</span></td>
  <td height="20"><input name="CCL" type="text" id="CCL"  value="<%=rs("CCL")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_SmallPACK"></span> <span class="red">*</span></td>
  <td height="20"><input name="SMALL_PACK" type="text" id="SMALL_PACK"  value="<%=rs("SMALL_PACK")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_BoxSize"></span> <span class="red">*</span></td>
  <td height="20"> <input name="Box_Size" type="text" id="Box_Size"  value="<%=rs("Box_Size")%>" size="8"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerPN"></span> <span class="red">*</span></td>
  <td height="20"><input name="Customer_PN" type="text" id="Customer_PN"  value="<%=rs("Customer_PN")%>" size="10"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerDefine"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_DEFINE" type="text" id="CUSTOMER_DEFINE"  value="<%=rs("CUSTOMER_DEFINE")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerLabel"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_LABEL" type="text" id="CUSTOMER_LABEL"  value="<%=rs("CUSTOMER_LABEL")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerDesc"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_DESC" type="text" id="CUSTOMER_DESC"  value="<%=rs("CUSTOMER_DESC")%>" size="50"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerPegapn"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_PEGAPN" type="text" id="CUSTOMER_PEGAPN"  value="<%=rs("CUSTOMER_PEGAPN")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerConfig"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_CONFIG" type="text" id="CUSTOMER_CONFIG"  value="<%=rs("CUSTOMER_CONFIG")%>" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="td_YNLittleLable"></span> <span class="red">*</span></td>
  <td height="20"> <select name="YNLittleLable" id="YNLittleLable" >
       <option value="NO" <%if rs("YESNOLITTLELABLE")="NO" then%>selected<%end if%>>NO</option>
      <option value="YES" <%if rs("YESNOLITTLELABLE")="YES" then%>selected<%end if%>>YES</option>
     
      
    </select></td>
</tr>
<tr>
  <td height="20"><span id="inner_Status"></span> <span class="red">*</span></td>
  <td height="20"> <select name="status" id="status" >
      <option>Status</option>
      <option value="ACTIVE" <%if rs("ITEM_STATUS")="ACTIVE" then%>selected<%end if%>>ACTIVE</option>
      <option value="INACTIVE" <%if rs("ITEM_STATUS")="INACTIVE" then%>selected<%end if%>>INACTIVE</option>
      <option value="OBSOLETE" <%if rs("ITEM_STATUS")="OBSOLETE" then%>selected<%end if%>>OBSOLETE</option>
      <option value="LTB" <%if rs("ITEM_STATUS")="LTB" then%>selected<%end if%>>LTB</option>
      <option value="PROTO TYPE" <%if rs("ITEM_STATUS")="PROTO TYPE" then%>selected<%end if%>>PROTO TYPE</option>
	  <option value="GREEN" <%if rs("ITEM_STATUS")="GREEN" then%>selected<%end if%>>GREEN</option>
	  <option value="ENTERING" <%if rs("ITEM_STATUS")="ENTERING" then%>selected<%end if%> >ENTERING</option>
    </select></td>
</tr>
<tr>
  <td height="20"><span id="td_LeadTime"></span></td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
    <input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="lead_time_unit" id="lead_time_unit">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
<tr>
  <td height="20"><span id="td_LeadTime2"></span></td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME2")),newtime)%>
    <input name="lead_time2" type="text" id="lead_time2" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="lead_time_unit2" id="lead_time_unit2">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="modelname" type="hidden" id="modelname" value="<%=rs("ITEM_NAME")%>">
      <input name="id" type="hidden" id="id" value="<%=id%>">
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