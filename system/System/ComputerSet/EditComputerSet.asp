<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->

<%
pagename="EditComputerSet.asp"
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query

computer=request("computer")
printer=request("printer")
line=request("line")
line=replace(line," ","")
PRODUCT=trim(request("PRODUCT"))

action=request("action")
if computer <> "" then
	strSql="select computer_name,printer_name,line_name,PRODUCT from computer_printer_mapping where computer_name='"&computer&"'"
	rs.open strSql,conn,1,3
	if not rs.eof then
		if action="update" then
			rs("printer_name")=printer
			rs("line_name")=line
			rs("PRODUCT")=PRODUCT
			rs.update
			word="Edit data successfully."
			action="location.href='"&beforepath&"'"
		else
			printer=rs("printer_name")
			line=rs("line_name")
			PRODUCT=rs("PRODUCT")
		end if		
	else
		word="Computer name of "&computer&" does not exist."
		action="location.href='"&beforepath&"'"
	end if
	rs.close		
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language="JavaScript">
if("<%=word%>" != ""){
	alert("<%=word%>");
	<%=action%>;
}
function  formcheck(){
	if(!form1.computer.value){
		alert("Computer Name cannot be blank.\n计算机名称不能为空.");
		form1.computer.focus();
		return false;
	}
}

</script>
</head>

<body onLoad="language(<%=session("language")%>);">

<form name="form1" method="post" action="" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
    <tr>
      <td width="120" height="20"><span id="td_Computer"></span>  <span class="red">*</span></td>
      <td  height="20"><%=computer%>
	  	<input name="computer" type="hidden" id="computer" value="<%=computer%>"> 
	  </td>     
    </tr>
    <tr> 
      <td height="20"><span id="td_Printer"></span></td>
      <td height="20"><input name="printer" type="text" id="printer"  value="<%=printer%>">	  </td>
    </tr>
    <tr>
      <td height="20"><span id="td_LineName"></span></td>
      <td height="20"><%= getLine("CHECKBOX",line,"","","") %>
	  </td>
    </tr>  
    <tr>
    <td height="20"><span id="inner_SubSeries"></span> <span class="red">*</span></td>
  <td height="20"><input name="PRODUCT" type="text" id="PRODUCT"  value="<%=PRODUCT%>" size="40" > 
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
	  	  <input name="action" type="hidden" id="action" value="update">
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