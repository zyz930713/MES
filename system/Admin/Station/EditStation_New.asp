<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetDefectCodeGroup.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from STATION_New where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Station/FormCheck.js" type="text/javascript"></script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">

<form action="/Admin/Station/EditStation_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
    </tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20"><span id="td_StationCode"></span> <span class="red">*</span> </td>
      <td height="20" class="red"><input name="stationnumber" type="text" id="stationnumber" value="<%=rs("STATION_NUMBER")%>" size="50"></td>
    </tr>
    <tr>
      <td width="183" height="20"><span id="td_Name"></span> <span class="red">*</span> </td>
      <td width="722" height="20" class="red">
        <div align="left">
          <input name="stationname" type="text" id="stationname" value="<%=rs("STATION_NAME")%>" size="50">
      </div></td>
    </tr>
    <tr>
      <td height="20"><span id="td_CHName"></span> <span class="red">*</span></td>
      <td height="20"><input name="stationchinesename" type="text" id="stationchinesename" value="<%=rs("STATION_CHINESE_NAME")%>" size="30"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value="">-- Select Section --</option>
          <%= getFactory("OPTION",rs("FACTORY_ID"),"","","") %>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Section"></span> <span class="red">*</span></td>
      <td height="20"><select name="section" id="section">
          <option value="">-- Select Section --</option>
          <%= getSection("OPTION",rs("SECTION_ID"),"","","") %>
        </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_TranType"></span>  <span class="red">*</span></td>
      <td height="20"><input name="transaction" type="radio" id="transaction" value="0" <%if rs("TRANSACTION_TYPE")="0" then%>checked<%end if%>>
        <span id="txt_Compulsory"></span>
        <input name="transaction" type="radio" value="1" id="transaction" <%if rs("TRANSACTION_TYPE")="1" then%>checked<%end if%>>
      <span id="txt_Optional"></span>
      <input name="transaction" type="radio" value="2" id="transaction" <%if rs("TRANSACTION_TYPE")="2" then%>checked<%end if%>>
<span id="txt_Conjunctive"></span></td>
    </tr>
   
    <tr>
      <td height="20"><span id="td_Wip"></span> </td>
      <td height="20"><input name="WIP" type="checkbox" id="WIP" value="1" <%if rs("WIP_REPORT_COLUMN")="1" then%>checked<%end if%>></td>
    </tr>
    <tr>
      <td height="20"><span id="td_WipSeq"></span></td>
      <td height="20"><select name="WIP_SEQ">
	  <option value=""></option>
	  <%for i=1 to 100%>
        <option value="<%=i%>" <%if rs("WIP_SEQUENCY")<>"" then 
		if cint(rs("WIP_SEQUENCY"))=i then%>selected<%end if
		end if%>><%=i%></option>
	  <%next%>
      </select></td>
    </tr>

    <tr>
      <td height="20"><span id="td_Output"></span></td>
      <td height="20"><input name="Output" type="checkbox" id="Output" value="1" <%if rs("Output_REPORT_COLUMN")="1" then%>checked<%end if%>></td>
    </tr>
    
    <tr>
      <td height="20"><span id="td_OutputSeq"></span></td>
      <td height="20"><select name="Output_SEQ">
          <option value=""></option>
          <%for i=1 to 100%>
          <option value="<%=i%>" <%if rs("OUTPUT_SEQUENCY")<>"" then 
		if cint(rs("OUTPUT_SEQUENCY"))=i then%>selected<%end if
		end if%>><%=i%></option>
          <%next%>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_IniQtyType"></span>  <span class="red">*</span> </td>
      <td height="20"><input name="quantity_type" type="radio" value="New" <%if rs("INITAIL_QUANTITY_TYPE")="New" then%>checked<%end if%>>
    <span id="txt_NewQty"></span>
      <input name="quantity_type" type="radio" value="Con" <%if rs("INITAIL_QUANTITY_TYPE")="Con" then%>checked<%end if%>>
    <span id="txt_ContQty"></span> </td> </td>
    </tr>
	<!--
   <tr>
      <td height="20">Tolerance</td>
	  <td><input name="Tolerance" type="text" id="Tolerance" value="<%=rs("TOLERANCE")%>">
	  	<select name="ToleranceUnit">
		  <option value="H" <% if rs("ToleranceUnit")="H" then response.write "selected" END IF%>>Hour</option>
          <option value="M" <% if rs("ToleranceUnit")="M" then response.write "selected" END IF%>>Minute</option>
		  <option value="S" <% if rs("ToleranceUnit")="S" then response.write "selected" END IF%>>Second</option>
	    </select>
	  </td>
    </tr>
	-->
    <tr>
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="actionscount" type="hidden" id="actionscount">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="btnOK" value="OK">
&nbsp;
        <input type="reset" name="btnReset" value="Reset">
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->