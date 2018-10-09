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
StationGroupName=request("txtstationgroupname")
StationChineseName=request("txtstationGroupChineseName")
STATION_GROUP_ID=request.QueryString("STATION_GROUP_ID")
if(STATION_GROUP_ID<>"") then
	set rsSTATION_GROUP=server.createobject("adodb.recordset")
	SQL="UPDATE STATION_GROUP SET STATUS='0',CREATE_TIME=SYSDATE,CREATE_USER_CODE='"+SESSION("code")+"' WHERE STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
	rsSTATION_GROUP.open SQL,conn,1,3
end if 
where=" where 1=1"
if(StationGroupName<>"") then
	where=where+" and STATION_GROUP_NAME like '"+StationGroupName+"%'" 
end if 

if(StationChineseName<>"") then
	where=where+" and STATION_GROUP_CHINESE_NAME like '"+StationChineseName+"%'" 
end if 

SQL="select * from STATION_GROUP "&where&" and STATUS='1' order by  STATION_GROUP_NAME"

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
<form action="/Admin/StationGroup/StationGroupList.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search Station Group</td>
  </tr>
  <tr>
    <td height="20">Station Group Name</td>
    <td><input name="txtstationgroupname" type="text" id="txtstationgroupname" value="<%=stationGroupName%>"></td>
    <td height="20">Station Group Chinese Name</td>
    <td><input name="txtstationGroupChineseName" type="text" id="txtstationGroupChineseName" value="<%=stationGroupChineseName%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">Browse Station Group list </td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/StationGroup/StationGroup.asp" target="main" class="white">Add a New Station Group</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center">Station Group Name</div></td>
  <td class="t-t-Borrow"><div align="center">Chinese Name </div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('StationGroup.asp?Action1=edit&STATION_GROUP_ID=<%=rs("STATION_GROUP_ID")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Station Group?')){form1.action='StationGroupList.asp?STATION_GROUP_ID=<%=rs("STATION_GROUP_ID")%>';form1.submit();}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	<%end if%>
    <td><div align="center"><%= rs("STATION_GROUP_NAME") %></div></td>
    <td><div align="center"><%= rs("STATION_GROUP_CHINESE_NAME") %></div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>  
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->