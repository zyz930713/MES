<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStationCodes.asp" -->
<!--#include virtual="/Functions/GetOperator.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
FactoryRight "S."
pagename="/Admin/Authority/EditStationAuthority.asp"
SQL="select S.*,F.FACTORY_NAME from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID where S.NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Authority/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/Authority/EditStationAuthority1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit Authority of a Station   </td>
</tr>
<tr>
  <td height="20">Factory</td>
  <td height="20"><%= rs("FACTORY_NAME") %>&nbsp;</td>
</tr>
<tr>
  <td width="183" height="20"><div align="left">Station Name 
    <input name="stationname" type="hidden" id="stationname" value="<%=rs("STATION_NAME")%>">
  </div></td>
    <td width="571" height="20">
      <div align="left"><%=rs("STATION_NAME")%> (<%=rs("STATION_CHINESE_NAME")%>)</div></td>
    </tr>
  <tr>
    <td height="20"><div align="left">Athorized Operators <span class="red">*</span></div></td>
    <td height="20"><div align="center">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"> <div align="center">Available Operators </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Operators</div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        
        <tr>
          <td rowspan="7"><select name="fromitem" size="15" multiple id="fromitem">
		  	<%FactoryRight "O."%>
			<%= getOperator(true,"OPTION",""," where O.CODE not in ('"&getStationCodes(rs("NID"),"','",idcount)&"') "&factorywhereoutsideand," order by O.CODE","","",opcount) %>
			<%a_opcount=opcount%>
			</select></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
          <td rowspan="7"><select name="toitem" size="15" multiple id="toitem">
			<%= getOperator(true,"OPTION",""," where O.CODE in ('"&getStationCodes(rs("NID"),"','",idcount)&"')"&factorywhereoutsideand," order by O.CODE","","",opcount) %>
			<%s_opcount=opcount%>
			</select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem)"></div></td>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem)"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem)"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
        </tr>
        <tr>
          <td>Total:<%= a_opcount%></td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </div></td>
    </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="operatorcount" type="hidden" id="operatorcount">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
