<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from LINE where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Line/FormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Admin/Line/Lan_Line.asp" -->
</head>

<body onLoad="language();">
<form action="/Admin/Line/EditLine1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
</tr>
<tr> 
	<td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
     <% =session("User") %></td>
</tr>
<tr>
  <td width="100" height="20"><div align="left"><span id="inner_LineName"></span><span class="red">*</span> </div></td>
    <td height="20" class="red">
      <div align="left">
        <input name="linename" type="text" id="linename" value="<%=rs("LINE_NAME")%>">
      </div></td>
    </tr>
<tr>
  <td height="20"><span id="inner_Factory"></span><span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value=""></option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
	</select>
  </td>
</tr>
<tr>
  <td height="20"><span id="inner_Section"></span> <span class="red">*</span></td>
  <td height="20"><select name="section" id="section">
    <option value=""></option>
    <%FactoryRight "S."%>
    <%= getSection("OPTION",rs("SECTION_ID"),factorywhereoutside,"","") %>
   </select> </td>
</tr>
<tr>
  <td height="20"><span id="inner_GroupLeader"></span> </td>
  <td height="20"><select name="leader" id="leader">
    <option value=""></option>
    <%FactoryRight "U."%>
    <%= getEngineer("OPTION",rs("LEADER"),factorywhereoutside," order by U.USER_CODE","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><span id="inner_Supervisor"></span> </td>
  <td height="20"><select name="supervisor" id="supervisor">
    <option value=""></option>
    <%FactoryRight "U."%>
    <%= getEngineer("OPTION",rs("SUPERVISOR"),factorywhereoutside," order by U.USER_CODE","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><span id="inner_MachineLabels"></span></td>
  <td height="20"><input name="labels" type="text" id="labels" value="<%=rs("MACHINE_LABELS")%>" size="80">
    <br>IT inputs the file name. Use , to seperate files. IT输入文件名，用 , 分隔不同文件。</td>
</tr>
<tr>
  <td height="20"><span id="inner_FactoryCode"></span></td>
  <td height="20"><input name="FACTORY_CODE" type="text" id="FACTORY_CODE" value="<%=rs("FACTORY_CODE")%>" size="80">
    <br>Use , to seperate string 。用 , 分隔多个字符。</td>
</tr>
<tr>
  <td height="20"><span id="inner_CodeLineName"></span></td>
  <td height="20"><input name="CODE_LINENAME" type="text" id="CODE_LINENAME" value="<%=rs("CODE_LINENAME")%>" size="80">
    <br>Use , to seperate string 。用 , 分隔多个字符。</td>
</tr>
<tr>
  <td height="20"><span id="inner_CodeName"></span></td>
  <td height="20"><input name="CODE_NAME" type="text" id="CODE_NAME" value="<%=rs("CODE_NAME")%>" size="80">
    <br>Use , to seperate string 。用 , 分隔多个字符。</td>
</tr>
<tr>
  <td height="20"><span id="inner_690CodeName"></span></td>
  <td height="20"><input name="CODE_NAME2" type="text" id="CODE_NAME2" value="<%=rs("CODE_NAME2")%>" size="80">
    <br>Use , to seperate string 。用 , 分隔多个字符。</td>
</tr>
<tr>
  <td height="20"><span id="inner_VersionNumber"></span></td>
  <td height="20"><input name="VERSION_NUMBER" type="text" id="VERSION_NUMBER" value="<%=rs("VERSION_NUMBER")%>" size="80">
    <br>Use , to seperate string 。用 , 分隔多个字符。</td>
</tr>
<tr>
  <td height="20"><span id="inner_CodeDate690"></span></td>
  <td height="20"><input name="Code_Date690" type="text" id="Code_Date690" value="<%=rs("Code_Date690")%>" size="80">
   </td>
</tr>
<tr>
  <td height="20"><span id="inner_CodeDate"></span></td>
  <td height="20"><input name="Code_Date" type="text" id="Code_Date" value="<%=rs("Code_Date")%>" size="80">
   </td>
</tr>
<tr>
    <td height="20"><span id="inner_Product"></span> <span class="red">*</span></td>
  <td height="20"><input name="PRODUCT" type="text" id="PRODUCT"  value="<%=PRODUCT%>" size="40" readonly> 
<select name="select11" onChange="(document.form1.PRODUCT.value=this.options[this.selectedIndex].value)">
<option >请选择</option>
<%
set rs_s=server.createobject("adodb.recordset")
rs_s.open "select SUBSERIES_NAME from SUBSERIES ",conn,1,3
%>
<%
while not rs_s.eof%>
<option value="<%=rs_s("SUBSERIES_NAME")%>"><%=rs_s("SUBSERIES_NAME")%></option>
<%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
</select>
      </td>
    </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
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