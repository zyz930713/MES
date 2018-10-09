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
SQL="select * from STATION where NID='"&id&"'"
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
</head>

<body>

<form action="/Admin/Station/EditStation1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="2" class="t-c-greenCopy">Edit a  Station</td>
    </tr>
    <tr>
      <td height="20">Station Number <span class="red">*</span> </td>
      <td height="20" class="red"><input name="stationnumber" type="text" id="stationnumber" value="<%=rs("STATION_NUMBER")%>" size="50"></td>
    </tr>
    <tr>
      <td width="183" height="20">Station Name <span class="red">*</span> </td>
      <td width="722" height="20" class="red">
        <div align="left">
          <input name="stationname" type="text" id="stationname" value="<%=rs("STATION_NAME")%>" size="50">
      </div></td>
    </tr>
    <tr>
      <td height="20">Station Chinese Name <span class="red">*</span></td>
      <td height="20"><input name="stationchinesename" type="text" id="stationchinesename" value="<%=rs("STATION_CHINESE_NAME")%>" size="30"></td>
    </tr>
    <tr>
      <td height="20">Belonged Factory <span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value="">-- Select Section --</option>
          <%= getFactory("OPTION",rs("FACTORY_ID"),"","","") %>
      </select></td>
    </tr>
    <tr>
      <td height="20">Section <span class="red">*</span></td>
      <td height="20"><select name="section" id="section">
          <option value="">-- Select Section --</option>
          <%= getSection("OPTION",rs("SECTION_ID"),"","","") %>
        </select>      </td>
    </tr>
    <tr>
      <td height="20">Transcation Type  <span class="red">*</span></td>
      <td height="20"><input name="transaction" type="radio" id="transaction" value="0" <%if rs("TRANSACTION_TYPE")="0" then%>checked<%end if%>>
        Compulsory
        <input name="transaction" type="radio" value="1" id="transaction" <%if rs("TRANSACTION_TYPE")="1" then%>checked<%end if%>>
      Optional
      <input name="transaction" type="radio" value="2" id="transaction" <%if rs("TRANSACTION_TYPE")="2" then%>checked<%end if%>>
Conjunctive</td>
    </tr>
    <tr>
      <td height="20">Transcation allowed to change </td>
      <td height="20"><input name="ischange" type="checkbox" id="ischange" value="1" <%if rs("TRANSACTION_CHANGE")="1" then%>checked<%end if%>></td>
    </tr>
    <tr>
      <td height="20">Include Actions  <span class="red">*</span></td>
      <td height="20"><div align="center">
          <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
            <tr>
              <td height="20" class="t-t-Borrow">
                <div align="center">Available Actions </div></td>
              <td height="20"><div align="center">&nbsp;</div></td>
              <td height="20" class="t-t-Borrow"><div align="center">Selected Actions </div></td>
              <td><div align="center">&nbsp;</div></td>
            </tr>
            <tr>
              <td rowspan="7"><select name="fromitem" size="8" multiple id="fromitem">
			  <%
			  FactoryRight "A."
			  if not isnull(rs("ACTIONS_INDEX")) then
			  ACTIONS_INDEX=" where A.NID not in('"&replace(rs("ACTIONS_INDEX"),",","','")&"')"&factorywhereoutsideandnull
			  else
			  ACTIONS_INDEX=factorywhereoutsideandnull
			  end if
			  %>
				<%= getAction(true,"OPTION","",ACTIONS_INDEX," order by A.ACTION_NAME","","") %>
              </select></td>
              <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
              <td rowspan="7"><select name="toitem" size="8" multiple id="toitem">
				<%if not isnull(rs("ACTIONS_INDEX")) then%><%= getAction(true,"OPTION",""," where A.NID in('"&replace(rs("ACTIONS_INDEX"),",","','")&"')","",rs("ACTIONS_INDEX"),"") %><%end if%>
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
          </table>
      </div></td>
    </tr>
    <tr>
      <td height="20">Stations to be entered Defect Code </td>
      <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"><div align="center">Available Stations </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td rowspan="7"><select name="fromitem1" size="8" multiple id="fromitem1">
            <%
			FactoryRight "S."
			if rs("STATION_ENTER_DEFECTCODE")<>"" then
			where=" where S.NID not in ('"&replace(rs("STATION_ENTER_DEFECTCODE"),",","','")&"')"&factorywhereoutsideand
			else
			where=factorywhereoutsideand
			end if%>
            <%= getStation(true,"OPTION","",where," order by S.STATION_NAME","","") %>
          </select></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem1,document.form1.toitem1);selectedcount()"></div></td>
          <td rowspan="7"><select name="toitem1" size="8" multiple id="toitem1">
              <%if rs("STATION_ENTER_DEFECTCODE")<>"" then
			where=" where S.NID in ('"&replace(rs("STATION_ENTER_DEFECTCODE"),",","','")&"')"
			%>
              <%= getStation(true,"OPTION","",where,"",rs("STATION_ENTER_DEFECTCODE"),"") %>
              <%end if%>
          </select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem1)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem1,document.form1.fromitem1);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem1)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem1,document.form1.toitem1);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem1)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem1,document.form1.fromitem1);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem1)"> </div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20">DefectCode Group </td>
      <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow"><div align="center">Available Stations </div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
            <td><div align="center">&nbsp;</div></td>
          </tr>
          <tr>
            <td rowspan="7"><select name="fromitem3" size="8" multiple id="fromitem3">
            <%
			FactoryRight "D."
			if rs("DEFECTCODE_GROUPS")<>"" then
			where=" where D.NID not in ('"&replace(rs("DEFECTCODE_GROUPS"),",","','")&"')"&factorywhereoutsideand
			else
			where=factorywhereoutsideand
			end if%>
            <%= getDefectCodeGroup("OPTION","",where,"","") %>
            </select></td>
            <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem3,document.form1.toitem3);selectedcount()"></div></td>
            <td rowspan="7"><select name="toitem3" size="8" multiple id="toitem3">
            <%if rs("DEFECTCODE_GROUPS")<>"" then
			where=" where D.NID in ('"&replace(rs("DEFECTCODE_GROUPS"),",","','")&"')"
			%>
            <%= getDefectCodeGroup("OPTION","",where,"","") %>
            <%end if%>
            </select></td>
            <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem3)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem3,document.form1.fromitem3);selectedcount()"></div></td>
            <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem3)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem3,document.form1.toitem3);selectedcount()"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem3)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem3,document.form1.fromitem3);selectedcount()"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem3)"> </div></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20">Is Column in WIP Report </td>
      <td height="20"><input name="WIP" type="checkbox" id="WIP" value="1" <%if rs("WIP_REPORT_COLUMN")="1" then%>checked<%end if%>></td>
    </tr>
    <tr>
      <td height="20">Sequency in WIP Report </td>
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
      <td height="20">Include stations for WIP Report </td>
      <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow">
            <div align="center">Available Stations </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td rowspan="7"><select name="fromitem2" size="10" multiple id="fromitem2">
            <%
			FactoryRight "S."
			if rs("WIP_INCLUDED_STATIONS")<>"" then
			where=" where S.NID not in ('"&replace(rs("WIP_INCLUDED_STATIONS"),",","','")&"') and NID<>'"&id&"'"&factorywhereoutsideand
			else
			where=" where S.NID<>'"&id&"'"&factorywhereoutsideand
			end if%>
              <%= getStation(true,"OPTION","",where," order by S.STATION_NAME","","") %>
          </select></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem2,document.form1.toitem2);selectedcount()"></div></td>
          <td rowspan="7"><select name="toitem2" size="10" multiple id="toitem2">
              <%if rs("WIP_INCLUDED_STATIONS")<>"" then
			where=" where S.NID in ('"&replace(rs("WIP_INCLUDED_STATIONS"),",","','")&"')"
			%>
              <%= getStation(true,"OPTION","",where,"",rs("WIP_INCLUDED_STATIONS"),"") %>
              <%end if%>
          </select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem2,document.form1.fromitem2);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem2,document.form1.toitem2);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem2,document.form1.fromitem2);selectedcount()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem2)"> </div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20">Is Column in Output Report </td>
      <td height="20"><input name="Output" type="checkbox" id="Output" value="1" <%if rs("Output_REPORT_COLUMN")="1" then%>checked<%end if%>></td>
    </tr>
    
    <tr>
      <td height="20">Sequency in Output Report </td>
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
      <td height="20">Intial Quantity Type  <span class="red">*</span> </td>
      <td height="20"><input name="quantity_type" type="radio" value="New" <%if rs("INITAIL_QUANTITY_TYPE")="New" then%>checked<%end if%>>
    Restart New Quanity
      <input name="quantity_type" type="radio" value="Con" <%if rs("INITAIL_QUANTITY_TYPE")="Con" then%>checked<%end if%>>
    Continue Previous Station </td>
    </tr>
    <tr>
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="actionscount" type="hidden" id="actionscount">
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->