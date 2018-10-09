<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetUserGroup.asp" -->
<!--#include virtual="/Functions/GetApprovalRole.asp" -->
<%
pagename="EditForm.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from FORM where NID='"&id&"'"
rs.open SQL,conn,1,3
if rs("APPROVAL_CODES")<>"" then
array_approval_code=split(rs("APPROVAL_CODES"),",")
else
array_approval_code=split(",,,,",",")
end if
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/System/Form/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form name="form1" method="post" action="/System/Form/EditForm1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Edit a Form </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">User:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="168" height="20">Form Name <span class="red">*</span></td>
      <td width="765" height="20"><input name="formname" type="text" id="formname" value="<%=rs("FORM_NAME")%>" size="50">	  </td>
    </tr>
    <tr>
      <td height="20">Form Chinese Name </td>
      <td height="20"><input name="formchinesename" type="text" id="formchinesename" value="<%=rs("FORM_CHINESE_NAME")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="20">Description</td>
      <td height="20"><input name="description" type="text" id="description" value="<%=rs("DESCRIPTION")%>" size="50">      </td>
    </tr>
    <tr>
      <td height="20">Apply Roles </td>
      <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"><div align="center">Available Roles </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Roles </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td height="20" rowspan="7"><select name="fromitem" size="10" multiple id="fromitem">
              <%if rs("APPLY_GROUP")<>"" then
			APPLY_GROUP=" where U.NID not in ('"&replace(rs("APPLY_GROUP"),",","','")&"')"
			end if%>
              <%= getUserGroup("OPTION","",APPLY_GROUP," order by GROUP_NAME","") %>
          </select></td>
          <td height="20"><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
          <td height="20" rowspan="7"><select name="toitem" size="10" multiple id="toitem">
              <%if rs("APPLY_GROUP")<>"" then
			APPLY_GROUP=" where U.NID in ('"&replace(rs("APPLY_GROUP"),",","','")&"')"
			%>
              <%= getUserGroup("OPTION","",APPLY_GROUP," order by GROUP_NAME","") %>
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
      <td height="20">Package</td>
      <td height="20"><input name="package" type="text" id="package" value="<%=rs("PACKAGE")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Control Type 1 </td>
      <td height="20" class="t-b-blue"><select name="paramtype1" id="paramtype1">
        <option>--Select Type--</option>
        <option value="text" <%if rs("PARAM_TYPE1")="text" then%>selected<%end if%>>Text</option>
        <option value="radio" <%if rs("PARAM_TYPE1")="radio" then%>selected<%end if%>>Radio</option>
		<option value="option" <%if rs("PARAM_TYPE1")="option" then%>selected<%end if%>>Option</option>
      </select></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Name 1 </td>
      <td height="20" class="t-b-blue"><input name="param1" type="text" id="param1" value="<%=rs("PARAM1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Chinese Name 1 </td>
      <td height="20" class="t-b-blue"><input name="cparam1" type="text" id="cparam1" value="<%=rs("PARAM_CHINESE1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Scripts 1 </td>
      <td height="20" class="t-b-blue"><input name="paramscripts1" type="text" id="paramscripts1" value="<%=rs("PARAM_SCRIPTS1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter  Button 1</td>
      <td height="20" class="t-b-blue"><input name="paramshowbutton1" type="text" id="paramshowbutton1" value="<%=rs("PARAM_SHOWBUTTON1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Chinese Button 1 </td>
      <td height="20" class="t-b-blue"><input name="paramchineseshowbutton1" type="text" id="paramchineseshowbutton1" value="<%=rs("PARAM_CHINESE_SHOWBUTTON1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Button Scripts 1 </td>
      <td height="20" class="t-b-blue"><input name="parambuttonscripts1" type="text" id="parambuttonscripts1" value="<%=rs("PARAM_BUTTON_SCRIPTS1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Title 1 </td>
      <td height="20" class="t-b-blue"><input name="paramtitle1" type="text" id="paramtitle1" value="<%=rs("PARAM_TITLE1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-blue">Parameter Chinese Title 1 </td>
      <td height="20" class="t-b-blue"><input name="paramchinesetitle1" type="text" id="paramchinesetitle1" value="<%=rs("PARAM_CHINESE_TITLE1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Control Type 2 </td>
      <td height="20" class="t-b-Yellow"><select name="paramtype2" id="paramtype2">
        <option>--Select Type--</option>
        <option value="text" <%if rs("PARAM_TYPE2")="text" then%>selected<%end if%>>Text</option>
        <option value="radio" <%if rs("PARAM_TYPE2")="radio" then%>selected<%end if%>>Radio</option>
		<option value="option" <%if rs("PARAM_TYPE2")="option" then%>selected<%end if%>>Option</option>
      </select></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Name 2 </td>
      <td height="20" class="t-b-Yellow"><input name="param2" type="text" id="param2" value="<%=rs("PARAM2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Chinese Name 2 </td>
      <td height="20" class="t-b-Yellow"><input name="cparam2" type="text" id="cparam2" value="<%=rs("PARAM_CHINESE2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Scripts 2 </td>
      <td height="20" class="t-b-Yellow"><input name="paramscripts2" type="text" id="paramscripts2" value="<%=rs("PARAM_SCRIPTS2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter  Button 2 </td>
      <td height="20" class="t-b-Yellow"><input name="paramshowbutton2" type="text" id="paramshowbutton2" value="<%=rs("PARAM_SHOWBUTTON2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Chinese Button 2 </td>
      <td height="20" class="t-b-Yellow"><input name="paramchineseshowbutton2" type="text" id="paramchineseshowbutton2" value="<%=rs("PARAM_CHINESE_SHOWBUTTON2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Button Scripts 2 </td>
      <td height="20" class="t-b-Yellow"><input name="parambuttonscripts2" type="text" id="parambuttonscripts2" value="<%=rs("PARAM_BUTTON_SCRIPTS2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Title 2 </td>
      <td height="20" class="t-b-Yellow"><input name="paramtitle2" type="text" id="paramtitle2" value="<%=rs("PARAM_TITLE2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-b-Yellow">Parameter Chinese Title 2 </td>
      <td height="20" class="t-b-Yellow"><input name="paramchinesetitle2" type="text" id="paramchinesetitle2" value="<%=rs("PARAM_CHINESE_TITLE2")%>" size="50"></td>
    </tr>
    
    <tr>
      <td height="20" class="t-t-FIN">Parameter Control Type 3 </td>
      <td height="20" class="t-t-FIN"><select name="paramtype3" id="paramtype3">
          <option>--Select Type--</option>
          <option value="text" <%if rs("PARAM_TYPE3")="text" then%>selected<%end if%>>Text</option>
          <option value="radio" <%if rs("PARAM_TYPE3")="radio" then%>selected<%end if%>>Radio</option>
		  <option value="option" <%if rs("PARAM_TYPE3")="option" then%>selected<%end if%>>Option</option>
      </select></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Name 3 </td>
      <td height="20" class="t-t-FIN"><input name="param3" type="text" id="param3" value="<%=rs("PARAM3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Chinese Name 3 </td>
      <td height="20" class="t-t-FIN"><input name="cparam3" type="text" id="cparam3" value="<%=rs("PARAM_CHINESE3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Scripts 3 </td>
      <td height="20" class="t-t-FIN"><input name="paramscripts3" type="text" id="paramscripts3" value="<%=rs("PARAM_SCRIPTS3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter  Button 3 </td>
      <td height="20" class="t-t-FIN"><input name="paramshowbutton3" type="text" id="paramshowbutton3" value="<%=rs("PARAM_SHOWBUTTON3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Chinese Button 3 </td>
      <td height="20" class="t-t-FIN"><input name="paramchineseshowbutton3" type="text" id="paramchineseshowbutton3" value="<%=rs("PARAM_CHINESE_SHOWBUTTON3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Button Scripts 3 </td>
      <td height="20" class="t-t-FIN"><input name="parambuttonscripts3" type="text" id="parambuttonscripts3" value="<%=rs("PARAM_BUTTON_SCRIPTS3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Title 3 </td>
      <td height="20" class="t-t-FIN"><input name="paramtitle3" type="text" id="paramtitle3" value="<%=rs("PARAM_TITLE3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-FIN">Parameter Chinese Title 3 </td>
      <td height="20" class="t-t-FIN"><input name="paramchinesetitle3" type="text" id="paramchinesetitle3" value="<%=rs("PARAM_CHINESE_TITLE3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Control Type 4 </td>
      <td height="20" class="t-t-Book"><select name="paramtype4" id="paramtype4">
          <option>--Select Type--</option>
          <option value="text" <%if rs("PARAM_TYPE4")="text" then%>selected<%end if%>>Text</option>
          <option value="radio" <%if rs("PARAM_TYPE4")="radio" then%>selected<%end if%>>Radio</option>
          <option value="option" <%if rs("PARAM_TYPE4")="option" then%>selected<%end if%>>Option</option>
      </select></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Name 4 </td>
      <td height="20" class="t-t-Book"><input name="param4" type="text" id="param4" value="<%=rs("PARAM4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Chinese Name 4 </td>
      <td height="20" class="t-t-Book"><input name="cparam4" type="text" id="cparam4" value="<%=rs("PARAM_CHINESE4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Scripts 4 </td>
      <td height="20" class="t-t-Book"><input name="paramscripts4" type="text" id="paramscripts4" value="<%=rs("PARAM_SCRIPTS4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter  Button 4 </td>
      <td height="20" class="t-t-Book"><input name="paramshowbutton4" type="text" id="paramshowbutton4" value="<%=rs("PARAM_SHOWBUTTON4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Chinese Button 4 </td>
      <td height="20" class="t-t-Book"><input name="paramchineseshowbutton4" type="text" id="paramchineseshowbutton4" value="<%=rs("PARAM_CHINESE_SHOWBUTTON4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Button Scripts 4 </td>
      <td height="20" class="t-t-Book"><input name="parambuttonscripts4" type="text" id="parambuttonscripts4" value="<%=rs("PARAM_BUTTON_SCRIPTS4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Title 4</td>
      <td height="20" class="t-t-Book"><input name="paramtitle4" type="text" id="paramtitle4" value="<%=rs("PARAM_TITLE4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Book">Parameter Chinese Title 4 </td>
      <td height="20" class="t-t-Book"><input name="paramchinesetitle4" type="text" id="paramchinesetitle4" value="<%=rs("PARAM_CHINESE_TITLE4")%>" size="50"></td>
    </tr>
    
    
    <tr>
      <td height="20"><p>Approve Flow </p></td>
      <td height="20">
	  <%for i=0 to 6%>
	  		  <select name="approve<%=i+1%>" id="approve<%=i+1%>">
        <option value="">-- Select --</option>
		<%=getApprovalRole("OPTION",trim(rs("APPROVE" & cstr(cint(i)+1))&""),null,null,null)%>
      </select>
	  <%if i<6 then%>
	  &gt;= 
	  <input name="interval<%=i+1%>" type="text" id="interval" value="<%=trim(rs("AMOUNT" & cstr(cint(i)+2))&"")%>" size="4">
	  -&gt;
	  <%end if
	next%></td>
    </tr>
    <tr>
      <td height="20">Alert Time </td>
      <td height="20">exceeds
        <select name="alert_time" id="alert_time">
            <option value="">-- Select --</option>
            <option value="60" <%if rs("ALERT_TIME")="60" then%>selected<%end if%>>&gt; 1 hour</option>
            <option value="120" <%if rs("ALERT_TIME")="120" then%>selected<%end if%>>&gt; 2 hours</option>
			<option value="180" <%if rs("ALERT_TIME")="180" then%>selected<%end if%>>&gt; 3 hous</option>
			<option value="240" <%if rs("ALERT_TIME")="240" then%>selected<%end if%>>&gt; 4 hours</option>
			<option value="1440" <%if rs("ALERT_TIME")="1440" then%>selected<%end if%>>&gt; 24 hours</option>
			<option value="2880" <%if rs("ALERT_TIME")="2880" then%>selected<%end if%>>&gt; 48 hours</option>
			<option value="4320" <%if rs("ALERT_TIME")="4320" then%>selected<%end if%>>&gt; 72 hours</option>
        </select>
        to
        <select name="alert_person" id="alert_person">
          <option value="">-- Select --</option>
          <option value="4" <%if rs("ALERT_PERSON")="4" then%>selected<%end if%>>Other Supervisors</option>
          <option value="5" <%if rs("ALERT_PERSON")="5" then%>selected<%end if%>>Manager</option>
        </select></td>
    </tr>
    <tr>
      <td height="20">Acting Person </td>
      <td height="20"><input name="actperson" type="checkbox" id="actperson" value="applicant" <%if instr(rs("ACT_PERSON"),"applicant")>0 then%>checked<%end if%>>
        Applicant
        <input name="actperson" type="checkbox" id="actperson" value="groupleader" <%if instr(rs("ACT_PERSON"),"groupleader")>0 then%>checked<%end if%>>
        Job's Group Leader
        <input name="actperson" type="checkbox" id="actperson" value="supervisor" <%if instr(rs("ACT_PERSON"),"supervisor")>0 then%>checked<%end if%>>
        Job's Supervisor </td>
    </tr>
    
    

    
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="Submit" value="Submit">
          &nbsp; 
          <input type="reset" name="Submit2" value="Reset">
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->