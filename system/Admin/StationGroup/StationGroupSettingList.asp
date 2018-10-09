<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
STATION_GROUP_ID=request("stationgroup")

MAPPING_ID=request.QueryString("MAPPING_ID")

if(MAPPING_ID<>"") then
	set rsSTATION_GROUP_UVS=server.createobject("adodb.recordset")
	SQL="DELETE STATION_GROUP_UVS_SETTING WHERE MAPPING_ID='"+MAPPING_ID+"'"
	rsSTATION_GROUP_UVS.open SQL,conn,1,3
	
	set rsSTATION_GROUP_BOM=server.createobject("adodb.recordset")
	SQL="DELETE STATION_GROUP_BOM_SETTING WHERE MAPPING_ID='"+MAPPING_ID+"'"
	rsSTATION_GROUP_BOM.open SQL,conn,1,3
	
	set rsSTATION_GROUP_STATION_MAPPING=server.createobject("adodb.recordset")
	SQL="DELETE STATION_GROUP_STATION_MAPPING WHERE MAPPING_ID='"+MAPPING_ID+"'"
	rsSTATION_GROUP_STATION_MAPPING.open SQL,conn,1,3
	
end if 

if(STATION_GROUP_ID<>"") then
	where=where+" and B.station_group_id = '"+STATION_GROUP_ID+"'" 
end if 


SQL="select distinct A.mapping_id, b.station_group_name, f.family_name, s.series_name, ss.subseries_name, a.model_name, a.uvs "
SQL=SQL+" from STATION_GROUP_UVS_SETTING A, station_group B,FAMILY F,series_new S, subseries SS, product_model PM"
SQL=SQL+" WHERE a.station_group_id= b.station_group_id"
SQL=SQL+" AND a.family_id=F.NID(+)"
SQL=SQL+" AND a.series_id=S.NID(+)"
SQL=SQL+" AND a.subseries_id=SS.NID(+)"
if(STATION_GROUP_ID<>"0")then
SQL=SQL+ where
end if 

SQL=SQL+" order by STATION_GROUP_NAME"

session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language_page()">
<form action="/Admin/StationGroup/StationGroupSettingList.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn">Search Station Group Setting</td>
  </tr>
  <tr>
    <td height="20">Station Group Name</td>
    <td>
		<%
			set rsSTATION_GROUP=server.createobject("adodb.recordset")
			SQL="SELECT * FROM STATION_GROUP WHERE STATUS='1' order by STATION_GROUP_NAME "
			rsSTATION_GROUP.open SQL,conn,1,3
			
		%>
          <select name="stationgroup" id="stationgroup" style="width:110px">
		  <option value="0">-- Select --</option>
		  	<%while not rsSTATION_GROUP.eof%>
			<option value="<%=rsSTATION_GROUP("STATION_GROUP_ID")%>" <%if STATION_GROUP_ID=rsSTATION_GROUP("STATION_GROUP_ID") then response.write "Selected" end if%>><%=rsSTATION_GROUP("STATION_GROUP_NAME")%></option>
			<%
				rsSTATION_GROUP.movenext
				wend
			%>
			 </select>  
	</td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Station Group Setting list </td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/StationGroup/StationGroupSetting.asp" target="main" class="white">Add a New Station Group Setting</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center">Station Group Name</div></td>
  <td class="t-t-Borrow"><div align="center">Family </div></td>
  <td class="t-t-Borrow"><div align="center">Series </div></td>
  <td class="t-t-Borrow"><div align="center">SubSeries </div></td>
  <td class="t-t-Borrow"><div align="center">Model </div></td>
  <td class="t-t-Borrow"><div align="center">UVS </div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('StationGroupSetting.asp?Action1=edit&MAPPING_ID=<%=rs("MAPPING_ID")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Station Group Setting?')){form1.action='StationGroupSettingList.asp?MAPPING_ID=<%=rs("MAPPING_ID")%>';form1.submit();}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	<%end if%>
    <td><div align="center"><%= rs("STATION_GROUP_NAME")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("FAMILY_NAME")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("SERIES_NAME")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("SUBSERIES_NAME")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("MODEL_NAME")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("UVS") %>&nbsp;</div></td>
</tr>
<%

rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>  
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->