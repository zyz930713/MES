<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Action/ActionCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetGroup.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetMachine.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
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
<script language="JavaScript" src="/Admin/Action/FormCheck.js" type="text/javascript"></script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onload="language(<%=session("language")%>);">
<form action="/Admin/Action/AddAction_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck_New()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
</tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
<tr>
  <td  height="20"><div align="left"><span id="td_Name"></span> <span class="red">*</span></div></td>
    <td  height="20" class="red">
      <div align="left">
        <input name="actionname" type="text" id="actionname" size="50">
      </div></td>
    </tr>
<tr>
  <td  height="20"><div align="left"><span id="td_ActionCode"></span> <span class="red">*</span></div></td>
    <td  height="20" class="red">
      <div align="left">
        <input name="actioncode" type="text" id="actioncode" size="50">
      </div></td>
    </tr>
	
	
  <tr>
    <td height="20"><span id="td_CHName"></span> <span class="red">*</span></td>
    <td height="20"><input name="actionchinesename" type="text" id="actionchinesename" size="50"></td>
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
    <td height="20"><div align="left"><span id="td_Purpose"></span> <span class="red">*</span></div></td>
    <td height="20"><select name="actionpurpose" id="actionpurpose">
      <option></option>
	  <%for i=1 to Ubound(ActionPurpose)
	  		response.Write("<option value='"&i&"'>"&ActionPurpose(i)&"</option>")
	  	next
	  %>      
      <option value="0">Other</option>
    </select></td>
  </tr>
  <tr>
    <td height="20"><span id="td_Position"></span>  <span class="red">*</span></td>
    <td height="20"><select name="position" id="position">
      <option value=""></option>
      <option value="0">Before</option>
      <option value="1">After</option>
        </select></td>
  </tr>
  <tr>
    <td height="20"><span id="td_Append"></span> </td>
    <td height="20"><input name="append" type="radio" value="0" checked> 
      NO
        <input name="append" type="radio" value="1"> 
      YES</td>
  </tr>
  <tr>
    <td height="20"><span id="td_Null"></span></td>
    <td height="20"><input name="null" type="radio" value="0" checked>
      NO
      <input name="null" type="radio" value="1">
      YES</td>
  </tr>
  <tr>
    <td height="20"><span id="td_HasLotProperty"></span></td>
    <td height="20"><input name="has_lot" type="checkbox" id="has_lot" value="1"> 
    If action is used to scan part number of material, is there relative action to scan lot property for this par of material?<br>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;若该步骤是扫描材料料号，是否需要扫描对应的批号属性？    
	</td>
  </tr>
  <tr>
    <td height="20"><div align="left"><span id="td_ActionType"></span> <span class="red">*</span></div></td>
    <td height="20"><input type="radio" name="actiontype" value="Key">
      Key
      in
        <input name="actiontype" type="radio" value="Scan" checked>
      Scan</td>
  </tr>
  <tr>
    <td height="20"><div align="left"><span id="td_ComponentType"></span></div></td>
    <td height="20">
        <div align="left">TEXT</div></td></tr>
  <tr>
    <td height="20"><span id="td_CellNo"></span></td>
    <td height="20"><select name="componentnumber" id="componentnumber">
        <%for i=1 to 10%>
        <option value="<%=i%>"><%=i%></option>
        <%next%>
    </select></td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="rcount" type="hidden" value="<%=rcount%>">
	  <input name="gcount" type="hidden" value="<%=gcount%>">
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
<!--#include virtual="/Functions/TableControl.asp" -->
