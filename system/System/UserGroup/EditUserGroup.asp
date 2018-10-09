<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<!--#include virtual="/Functions/GetUser.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from USER_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/System/UserGroup/FormCheck.js" type="text/javascript"></script>
</head>

<body>

<form action="/System/UserGroup/EditUserGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a User Group  </td>
</tr>
<tr>
  <td width="135" height="20"><div align="left">Group Name <span class="red">*</span> </div></td>
  <td width="533" height="20" class="red">
    <div align="left">
      <input name="groupname" type="text" id="groupname" value="<%=rs("GROUP_NAME")%>" size="50">
  </div></td>
</tr>
<tr>
  <td height="20">Group Chinese Name <span class="red">*</span></td>
  <td height="20"><input name="groupchinesename" type="text" id="groupchinesename" value="<%=rs("GROUP_CHINESE_NAME")%>" size="50"></td>
</tr>
<tr>
  <td height="20">Belonged Factory <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value="">-- Select Factory --</option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20">Interface Laguage  <span class="red">*</span></td>
  <td height="20"><select name="language" id="language">
    <option>-- Select Language --</option>
    <option value="0" <%if cint(rs("LANGUAGE"))=0 then%>selected<%end if%>>English</option>
    <option value="1" <%if cint(rs("LANGUAGE"))=1 then%>selected<%end if%>>Chinese</option>
  </select>
  </td>
</tr>
<tr>
  <td height="20">Roles</td>
  <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center">Available Roles </div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20" class="t-t-Borrow"><div align="center">Selected Roles </div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20" rowspan="7"><select name="fromitem" size="10" multiple id="fromitem">
	  <%if rs("ROLES_ID")<>"" then
	  ROLES_ID=" where ROLE_NAME not in ('"&replace(rs("ROLES_ID"),",","','")&"')"
	  end if%>
      <%= getEngineerRole(true,"OPTION","",ROLES_ID," order by ROLE_NAME","","") %>
      </select></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
      <td height="20" rowspan="7"><select name="toitem" size="10" multiple id="toitem">
	  <%if rs("ROLES_ID")<>"" then
	  ROLES_ID=" where ROLE_NAME in ('"&replace(rs("ROLES_ID"),",","','")&"')"
	  %>
      <%= getEngineerRole(true,"OPTION","",ROLES_ID," order by ROLE_NAME","","") %>
	  <%end if%>
      </select></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20">Members</td>
  <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center">Available Roles </div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20" class="t-t-Borrow"><div align="center">Selected Roles </div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20" rowspan="7"><select name="fromitem2" size="10" multiple id="fromitem2">
	  <%if rs("GROUP_MEMBERS")<>"" then
	  USERS_ID=" where USER_CODE not in ('"&replace(rs("GROUP_MEMBERS"),",","','")&"')"
	  end if%>
      <%= getUser(true,"USER_GROUP_OPTION","",USERS_ID," order by USER_CODE","","") %>
      </select></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem2,document.form1.toitem2)"></div></td>
      <td height="20" rowspan="7"><select name="toitem2" size="10" multiple id="toitem2">
	  <%if rs("GROUP_MEMBERS")<>"" then
	  USERS_ID=" where USER_CODE in ('"&replace(rs("GROUP_MEMBERS"),",","','")&"')"%>
	  <%= getUser(true,"USER_GROUP_OPTION","",USERS_ID," order by USER_CODE","","") %>
	  <%end if%>
      </select></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem2)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem2,document.form1.fromitem2)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem2)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem2,document.form1.toitem2)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem2)"> </div></td>
    </tr>
    <tr>
      <td height="20"><div align="center">&nbsp;</div></td>
      <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
    <tr>
      <td height="20"><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem2,document.form1.fromitem2)"></div></td>
      <td height="20"><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem2)"> </div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20">Apply roles to member </td>
  <td height="20"><input name="apply" type="checkbox" id="apply" value="1"></td>
</tr>

  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
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
