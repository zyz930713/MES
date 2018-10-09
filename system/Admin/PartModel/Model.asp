<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Model/ModelCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
modelname=request("modelname")
modelstatus=request("modelstatus")
factory=request.QueryString("factory")
where=""
if modelname<>"" then
where=where&" and lower(segment1) like '%"&lcase(modelname)&"%'"
end if
if modelstatus<>"" and modelstatus<>"all" then
where=where&" and lower(inventory_item_status_code)='"&lcase(modelstatus)&"'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and P.FACTORY_ID is null"
else
	select case factory
	case "FA00000001"
	where=where&" and wip_supply_subinventory='C10 DELTEK'"
	case "FA00000002"
	where=where&" and wip_supply_subinventory='C10 ASSY'"
	case "FA00000003"
	where=where&" and wip_supply_subinventory='C10 VALADD'"
	end select
end if

pagename="/Admin/Model/Model.asp"
pagepara="&modelname="&modelname&"&modelstatus="&modelstatus&"&factory="&factory
ModelFactoryRight
SQLPR="select MODEL_NUMBER,INVENTORY_ITEM_STATUS_CODE,wip_supply_subinventory from tbl_System_Items where segment1 like '%-_____-%' or segment1 like 'SR%'"&where&modelfactoryand&" order by segment1"
'response.Write(SQLPR)
rsPR.open SQLPR,connPR,1,3
%>
<!--#include virtual="/Components/PageSelect_PR.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css"></head>

<body>
<form action="/Admin/Model/Model.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn">Search Model </td>
  </tr>
  <tr>
    <td width="11%" height="20">Model<strong> </strong> Name </td>
    <td width="18%"><input name="modelname" type="text" id="modelname" value="<%=modelname%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursPRor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy">Browse Model List </td>
</tr>
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="5"><!--#include virtual="/Components/PageSplit_PR.asp" --></td>
  </tr>
<tr>
  <td width="54" height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <%end if%>
  <td width="241" height="20" class="t-t-Borrow"><div align="center">Model Name </div></td>
  <td width="235" class="t-t-Borrow"><div align="center">Oranization</div></td>
  <td width="235" class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?modelname=<%=modelname%>&factory='+this.options[this.selectedIndex].value+'&modelstatus=<%=modelstatus%>'">
      <option value="all">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td width="235" class="t-t-Borrow"><div align="center">
    <select name="modelstatus" id="modelstatus" onChange="location.href='<%=pagename%>?modelname=<%=modelname%>&factory=<%=factory%>&modelstatus='+this.options[this.selectedIndex].value">
      <option>Status</option>
      <option value="all">All</option>
      <option value="ACTIVE" <%if modelstatus="ACTIVE" then%>selected<%end if%>>ACTIVE</option>
      <option value="INACTIVE" <%if modelstatus="INACTIVE" then%>selected<%end if%>>INACTIVE</option>
      <option value="OBSOLETE" <%if modelstatus="OBSOLETE" then%>selected<%end if%>>OBSOLETE</option>
      <option value="LTB" <%if modelstatus="LTB" then%>selected<%end if%>>LTB</option>
      <option value="PROTO TYPE" <%if modelstatus="PROTO TYPE" then%>selected<%end if%>>PROTO TYPE</option>
	  <option value="GREEN" <%if modelstatus="GREEN" then%>selected<%end if%>>GREEN</option>
	  <option value="ENTERING">ENTERING</option>
    </select>
  </div></td>
</tr>
<%
i=1
if not rsPR.eof then
rsPR.absolutepage=cint(session("strpagenum"))
while not rsPR.eof and i<=rsPR.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <%end if%>
    <td height="20"><div align="center"><%= rsPR("segment1") %></div></td>
    <td><div align="center"><%= rsPR("organization_id")%></div></td>
    <td><div align="center"><%= rsPR("wip_supply_subinventory")%>&nbsp;</div></td>
    <td><div align="center"><%= rsPR("inventory_item_status_code")%></div></td>
</tr>
<%
i=i+1
rsPR.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="5">No Records&nbsp;</td>
  </tr>
<%end if
rsPR.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->