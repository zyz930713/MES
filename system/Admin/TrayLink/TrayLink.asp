<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TrayLink/TrayLinkCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Language/Admin/TrayLink/Lan_TrayLink.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
partnumber=request("partnumber")
stationname=request("stationname")

if partnumber<>"" then
	where=where&" and pa.part_number like '"&partnumber&"%'"
end if
if stationname<>"" then
	where=where&" and pa.station_id ='"&stationname&"'"
end if

pagename="/Admin/TrayLink/TrayLink.asp"
SQL="select pa.part_number,pa.station_id,pa.station_seq,pa.tray_type,pa.tray_size,st.station_name,st.station_chinese_name from part_tray pa left join station_new st on pa.station_id=st.nid where 1=1 "&where&"order by pa.station_id,pa.station_seq"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>
<body onLoad="language();language_page();">
<form action="/Admin/TrayLink/TrayLink.asp" method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
    </tr>
    <tr><td><table ><tr>
      <td height="20"><span id="inner_PartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>" ></td>
      <td height="20"><span id="inner_StationName"></span></td>
      <td>
	  <select name="stationname" id="stationname">
	  	<option value=""></option>
	  	<%=getStation_New(true,"OPTION",stationname,"","order by station_name","","") %>
	  </select>
	  </td>
      <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr></table></td></tr>
  </table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_TrayLinkList"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
            <% =session("User") %></td>
          <td width="50%"><div align="right">
              <%if admin=true then%>
              <a href="/Admin/TrayLink/AddTrayLink.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_NewTrayLink"></span></a>
              <%end if%>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_No"></span></div></td>
    <%if admin=true then%>
    <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
    <%end if%>
    <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber1"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StationName1"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StationChineseName"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StationSequence"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_TrayType"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_TraySize"></span></div></td>
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
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditTrayLink.asp?partnumber=<%=rs("part_number")%>&stationid=<%=rs("station_id")%>&traytype=<%=rs("tray_type")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
	<td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('DeleteTrayLink.asp?partnumber=<%=rs("part_number")%>&stationid=<%=rs("station_id")%>&traytype=<%=rs("tray_type")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
    <%end if%>
    <td><div align="center"><%= rs("part_number") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("station_name") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("station_chinese_name") %>&nbsp;</div></td>
	<td><div align="center"><%= rs("station_seq") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("tray_type") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("tray_size") %>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center"><span id="inner_NoRecord"></span></div></td>
  </tr>
  <%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
