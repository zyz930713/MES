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

<form action="/Admin/Station/AddStation_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>	
    <tr>
      <td width="20%" height="20"><span id="td_StationCode"></span>  <span class="red">*</span> </td>
      <td width="80%" height="20" class="red"><input name="stationnumber" type="text" id="stationnumber" size="30"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Name"></span> <span class="red">*</span> </td>
      <td height="20" class="red">
        <div align="left">
          <input name="stationname" type="text" id="stationname" size="30">
      </div></td>
    </tr>
    <tr>
      <td height="20"><span id="td_CHName"></span> <span class="red">*</span></td>
      <td height="20"><input name="stationchinesename" type="text" id="stationchinesename" size="30"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value=""></option>
		  <%FactoryRight ""%>
          <%= getFactory("OPTION","",factorywhereinside,"","") %>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Section"></span> <span class="red">*</span></td>
      <td height="20"><select name="section" id="section">
          <option value=""></option>
		  <%FactoryRight "S."%>
          <%= getSection("OPTION","",factorywhereoutside,"","") %>
        </select>      
		</td>
    </tr>
    <tr>
      <td height="20"><span id="td_TranType"></span> <span class="red">*</span></td>
      <td height="20"><input name="transaction" type="radio" id="transaction" value="0" checked>
<span id="txt_Compulsory"></span>
  <input name="transaction" type="radio" value="1" id="transaction">
<span id="txt_Optional"></span>
<input name="transaction" type="radio" value="2" id="transaction"> 
<span id="txt_Conjunctive"></span>
</td>
    </tr>    
    <tr>
      <td height="20"><span id="td_Wip"></span> </td>
      <td height="20"><input name="WIP" type="checkbox" id="WIP" value="1"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_WipSeq"></span> </td>
      <td height="20"><select name="WIP_SEQ">
          <option value=""></option>
          <%for i=1 to 20%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
      </select></td>
    </tr>
   
    <tr>
      <td height="20"><span id="td_Output"></span></td>
      <td height="20"><input name="Output" type="checkbox" id="Output" value="1"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_OutputSeq"></span> </td>
      <td height="20"><select name="Output_SEQ">
          <option value=""></option>
          <%for i=1 to 20%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_IniQtyType"></span> <span class="red">*</span> </td>
      <td height="20"><input type="radio" name="quantity_type" value="New">
        <span id="txt_NewQty"></span>
          <input name="quantity_type" type="radio" value="Con" checked>
        <span id="txt_ContQty"></span> </td>
    </tr>
	<!--
    <tr>	
      <td height="20">Tolerance</td>
	  <td><input name="Tolerance" type="text" id="Tolerance" value="1">
	  	<select name="ToleranceUnit">
		  <option value="H">Hour</option>
          <option value="M">Minute</option>
		  <option value="S">Second</option>
	    </select>
	  </td>
    </tr>
	-->
    <tr>
      <td height="20" colspan="2"><div align="center">
          <input name="actionscount" type="hidden" id="actionscount" value="0">
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