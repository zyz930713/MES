<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<!--#include virtual="/Functions/GetApprovalRole.asp" -->
<!--#include virtual="/Functions/GetEvent.asp" -->
<%
pagename="NewEngineer.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from USERS where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/System/Engineer/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onload="language(<%=session("language")%>);">
<form name="form1" method="post" action="/System/Engineer/EditEngineer1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
    <tr>
      <td width="14%" height="20"><span id="td_Code"></span> <span class="red">*</span></td>
      <td width="86%" height="20"><input name="code" type="text" id="code" value="<%=rs("USER_CODE")%>"></td>
    </tr>
    <tr> 
      <td height="20"><span id="td_Name"></span> <span class="red">*</span></td>
      <td height="20"><input name="name" type="text" id="name" value="<%=rs("USER_NAME")%>">	  </td>
    </tr>
    <tr>
      <td height="20"><span id="td_CHName"></span> <span class="red">*</span> </td>
      <td height="20"><input name="chinesename" type="text" id="chinesename" value="<%=rs("USER_CHINESE_NAME")%>"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_NTAccount"></span><span class="red">*</span></td>
      <td height="20"><input name="NT" type="text" id="NT" value="<%=rs("NT_ACCOUNT")%>"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="td_Manager"></span></td>
      <td height="20"><select name="manager" id="manager">
          <option value=""></option>
          <%= getEngineer("OPTION",rs("manager"),""," order by U.USER_CODE","") %>
      </select></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
      <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow"><div align="center"><span id="td_ValidFactory"></span></div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SeltFactory"></span></div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
          </tr>
          <tr>
            <td height="20" rowspan="4"><select name="factoryfrom" size="5" multiple id="factoryfrom">
                <%if rs("FACTORY_ID")<>"" then
			FACTORY_ID=" where instr('"&rs("FACTORY_ID")&"',NID)<=0"
			end if%>
                <%= getFactory("OPTION",null,FACTORY_ID," order by FACTORY_NAME",null) %>
            </select></td>
            <td height="20"><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.factoryfrom,document.form1.factoryto)"></div></td>
            <td height="20" rowspan="4"><select name="factoryto" size="5" multiple id="factoryto">
                <%if rs("FACTORY_ID")<>"" then
			FACTORY_ID=" where instr('"&rs("FACTORY_ID")&"',NID)>0"
			%>
                <%= getFactory("OPTION",null,FACTORY_ID," order by FACTORY_NAME",null) %>
                <%end if%>
            </select></td>
            <td height="20"><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.factoryto)"> </div></td>
          </tr>
          <tr>
            <td height="20"><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.factoryto,document.form1.factoryfrom)"></div></td>
            <td height="20"><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.factoryto)"> </div></td>
          </tr>
          
          <tr>
            <td height="20"><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.factoryfrom,document.form1.factoryto)"></div></td>
            <td height="20"><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.factoryto)"> </div></td>
          </tr>
          
          <tr>
            <td height="20"><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.factoryto,document.form1.factoryfrom)"></div></td>
            <td height="20"><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.factoryto)"> </div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Language"></span></td>
      <td height="20"><select name="language" id="language">
        <option value="0" <%if rs("LANGUAGE")="0" then%>selected<%end if%>>English</option>
        <option value="1" <%if rs("LANGUAGE")="1" then%>selected<%end if%>>Chinese</option>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Role"></span></td>
      <td height="20">
	    <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow">
              <div align="center"><span id="td_ValidRole"></span></div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SeltRole"></span></div></td>
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
      <td height="20"><span id="td_Email"></span></td>
      <td height="20"><input name="email" type="text" id="email" value="<%=rs("EMAIL")%>" size="50">      </td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="rolescount" type="hidden" id="rolescount">
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