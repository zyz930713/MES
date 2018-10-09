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
</head>

<body>
<form action="/Admin/Action/AddAction1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Add a New Action  </td>
</tr>
<tr>
  <td width="113" height="20"><div align="left">Action Name <span class="red">*</span></div></td>
    <td width="641" height="20" class="red">
      <div align="left">
        <input name="actionname" type="text" id="actionname" size="50">
      </div></td>
    </tr>
  <tr>
    <td height="20">Action Chinese Name <span class="red">*</span></td>
    <td height="20"><input name="actionchinesename" type="text" id="actionchinesename" size="50"></td>
  </tr>
  <tr>
    <td height="20">Belonged Factory <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
		<%FactoryRight ""%>
        <%= getFactory("OPTION","",factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20">Belonged Station</td>
    <td height="20">
      <div align="left">
        <select name="station" id="station">
          <option value="">-- None --</option>
		  <%FactoryRight ""%>
          <%=getStation(true,"OPTION","",factorywhereoutside," order by STATION_NAME","","")%>
        </select>
    </div></td>
  </tr>
  <tr>
    <td height="20"><div align="left">Action Purpose  <span class="red">*</span></div></td>
    <td height="20"><select name="actionpurpose" id="actionpurpose">
      <option>-- Select Purpose --</option>
      <option value="1">Machine Code</option>
      <option value="2">Material Part Number</option>
      <option value="3">Material Lot Number</option>
      <option value="4">Material Quantity</option>
      <option value="5">Job Quantity</option>
	  <option value="6">Rework Quantity</option>
      <option value="0">Other</option>
    </select></td>
  </tr>
  <tr>
    <td height="20">Position  <span class="red">*</span></td>
    <td height="20"><select name="position" id="position">
      <option>-- Select Position --</option>
      <option value="0">Before</option>
      <option value="1">After</option>
        </select></td>
  </tr>
  <tr>
    <td height="20">Append Allowed </td>
    <td height="20"><input name="append" type="radio" value="0" checked> 
      NO
        <input name="append" type="radio" value="1"> 
      YES</td>
  </tr>
  <tr>
    <td height="20">Null Allowed </td>
    <td height="20"><input name="null" type="radio" value="0" checked>
      NO
      <input name="null" type="radio" value="1">
      YES</td>
  </tr>
  <tr>
    <td height="20">Has Lot Property </td>
    <td height="20"><input name="has_lot" type="checkbox" id="has_lot" value="1"> 
      If action is used to scan part number of material, is there relative action to scan lot property for this par of material?</td>
  </tr>
  <tr>
    <td height="20"><div align="left">Valid Machine </div></td>
    <td height="20"><input name="machine" type="text" id="machine" size="20"> 
      None means irrelevant; % means any character </td>
  </tr>
  <tr>
    <td height="20"><div align="left">Valid Value </div></td>
    <td height="20"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top"><table width="100%" border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
            <tr>
              <td height="20" class="t-t-Borrow"><div align="center">Part Number</div></td>
              <td height="20" class="t-t-Borrow"><div align="center">Material</div></td>
              <td height="20" class="t-t-Borrow"><div align="center">Machine</div></td>
            </tr>
			<%FactoryRight "P."%>
            <%=getPart(true,"ACTION_TABLE_VALUE",""," where P.ROUTINE_TYPE=0 and P.STATUS=1 "&factorywhereoutsideand," order by P.SECTION_ID","","",rcount,id,"",tabi,"","")%>
          </table></td>
          <td valign="top"><table width="100%" border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
              <tr>
                <td height="20" class="t-t-Borrow"><div align="center">Group Name</div></td>
                <td height="20" class="t-t-Borrow"><div align="center">Material</div></td>
                <td height="20" class="t-t-Borrow"><div align="center">Machine</div></td>
              </tr>
			  <%FactoryRight "SG."%>
              <%=getGroup(true,"TABLE_VALUE",""," where SG.STATUS=1 "&factorywhereoutsideand,"","","",gcount,id)%>
          </table></td>
        </tr>
        <tr>
          <td colspan="2">None means irrelevant; * means any simgle character; % means any characters; use &quot;,&quot; to separate multi-format. If  Valid Machine is not none, it overrade Valid Machine settings. </td>
        </tr>
      </table>      </td>
  </tr>
  <tr>
    <td height="20"><div align="left">Related Action </div></td>
    <td height="20"><select name="relatedaction" id="relatedaction">
	<option value="">-- None --</option>
      <%= getAction(true,"OPTION","",""," order by ACTION_NAME","","")%>
    </select>      <br> 
      If  action is used to key in quantity of certain material, select which action defines part number of this material. None means irrelevant</td>
  </tr>
  <tr>
    <td height="20"><div align="left">Max Quantity </div></td>
    <td height="20"><input name="max" type="text" id="max" value="0" size="10"> 
      0 means unlimited or irrelevant</td>
  </tr>
  <tr>
    <td height="20"><div align="left">Min Quantity </div></td>
    <td height="20"><input name="min" type="text" id="min" value="0" size="10">
      0 means unlimited or irrelevant</td>
  </tr>
  <tr>
    <td height="20"><div align="left">Action Type  <span class="red">*</span></div></td>
    <td height="20"><input type="radio" name="actiontype" value="Key">
      Key
      in
        <input name="actiontype" type="radio" value="Scan" checked>
      Scan</td>
  </tr>
  <tr>
    <td height="20"><div align="left">Component Type <span class="red">*</span> </div></td>
    <td height="20">
        <div align="left">
          <select name="componenttype" id="componenttype">
            <option>-- Select Type --</option>
            <option value="TEXT">TEXT</option>
          </select>
      </div></td></tr>
  <tr>
    <td height="20">Cells Number</td>
    <td height="20"><select name="componentnumber" id="componentnumber">
        <option value="">-- Select Number --</option>
        <%for i=5 to 20%>
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
      <input type="submit" name="Submit" value="Save">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/Functions/TableControl.asp" -->
