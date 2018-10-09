<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Action/ActionCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Admin/Action/Action.asp"
actionname=request("actionname")
chineseactionname=request("chineseactionname")
thisstatus=request.QueryString("status")
station=request("station")
factory=request.QueryString("factory")
purpose=request.QueryString("purpose")
actionNID=request("actionNID")
where=""
if actionname<>"" then
where=where&" and lower(A.ACTION_NAME) like '%"&lcase(actionname)&"%'"
end if
if chineseactionname<>"" then
where=where&" and lower(A.ACTION_CHINESE_NAME) like '%"&lcase(chineseactionname)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and A.STATUS="&thisstatus
end if
if station="" or station="all" then
where=where&""
elseif station="null" then
where=where&" and A.STATION_ID is null"
else
where=where&" and A.STATION_ID='"&station&"'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and A.FACTORY_ID is null"
else
where=where&" and A.FACTORY_ID='"&factory&"'"
end if
if purpose="" or purpose="all" then
where=where&""
else
where=where&" and A.ACTION_PURPOSE="&purpose
end if
if actionNID<>"" then
	where=where&" and A.NID='"&actionNID&"' "
end if
pagepara="&status="&thisstatus&"&station="&station&"&factory="&factory&"&purpose="&purpose
FactoryRight "A."
SQL="select A.NID,A.NID,A.ACTION_NAME,A.ACTION_CHINESE_NAME,A.STATUS,A.STATION_ID,A.ACTION_PURPOSE,A.STATION_POSITION,A.APPEND_ALLOW,A.NULL_ALLOW,A.VALID_MACHINE,A.RELATED_ACTION_ID,A.MAX_QUANTITY,A.MIN_QUANTITY,A.ACTION_TYPE,A.ELEMENT_TYPE,A.ELEMENT_NUMBER,A.WITH_LOT,ST.STATION_NAME,ST.STATION_CHINESE_NAME,F.FACTORY_NAME from ACTION A left join STATION ST on A.STATION_ID=ST.NID left join Factory F on A.FACTORY_ID=F.NID where 1=1 "&where&factorywhereoutsideand&" Order by A.STATION_ID,A.ACTION_PURPOSE"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<form action="/Admin/Action/Action.asp" method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="7" class="t-b-midautumn">Search Action </td>
    </tr>
    <tr>
	  <td>Action ID</td>
      <td><input name="actionNID" type="text" id="actionNID" value="<%=actionNID%>"></td>	
      <td height="20">Action Name</td>
      <td height="20"><input name="actionname" type="text" id="actionname" value="<%=actionname%>"></td>
      <td>Action Chinese Name</td>
      <td><input name="chineseactionname" type="text" id="chineseactionname" value="<%=chineseactionname%>"></td>
      
    </tr>
    <tr>
      <td>Station</td>
      <td colspan="4"><select name="station" id="station">
          <option value="">Station</option>
          <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
          <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
          <%FactoryRight "S."%>
          <%=getStation(true,"OPTION",station,factorywhereoutside," order by STATION_NAME","","")%>
      </select></td>
      <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
  </table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="21" class="t-c-greenCopy">Browse Action List </td>
</tr>
<tr>
  <td height="20" colspan="21" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Action/AddAction.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Action</a><%end if%></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="21"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">NID</div></td>

  <td class="t-t-Borrow"><div align="center">
    <select name="status" id="status" onChange="location.href='<%=pagename%>?status='+this.options[this.selectedIndex].value+'&factory=<%=factory%>&purpose=<%=purpose%>'">
      <option value="">Status</option>
      <option value="all" <%if thisstatus="all" then%>selected<%end if%>>All</option>
	  <option value="1" <%if thisstatus="1" then%>selected<%end if%>>Enabled</option>
      <option value="0" <%if thisstatus="0" then%>selected<%end if%>>Disabled</option>
    </select>
  </div></td>
  
  <td class="t-t-Borrow"><div align="center">Station</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Action Name </div></td>
  <td class="t-t-Borrow"><div align="center"> Chinese Name </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?&status=<%=thisstatus%>&factory='+this.options[this.selectedIndex].value+'&purpose=<%=purpose%>'">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="purpose" id="purpose" onChange="location.href='<%=pagename%>?status=<%=thisstatus%>&factory=<%=factory%>&purpose='+this.options[this.selectedIndex].value">
      <option>Purpose</option>
      <option value="all" <%if purpose="all" then%>selected<%end if%>>All</option>
      <option value="1" <%if purpose="1" then%>selected<%end if%>>Machine Code</option>
      <option value="2" <%if purpose="2" then%>selected<%end if%>>Material Part</option>
      <option value="3" <%if purpose="3" then%>selected<%end if%>>Material Lot</option>
      <option value="4" <%if purpose="4" then%>selected<%end if%>>Material Quantity</option>
      <option value="5" <%if purpose="5" then%>selected<%end if%>>Job Quantity</option>
      <option value="0" <%if purpose="0" then%>selected<%end if%>>Other</option>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">Position</div></td>
  <td class="t-t-Borrow"><div align="center">Null</div></td>
  <td class="t-t-Borrow"><div align="center">Append</div></td>
  <td class="t-t-Borrow"><div align="center">Lot</div></td>
  <td class="t-t-Borrow"><div align="center">Valid Machine </div></td>
  <td class="t-t-Borrow"><div align="center">Related Action </div></td>
  <td class="t-t-Borrow"><div align="center">Max Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Min Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Action Type </div></td>
  <td class="t-t-Borrow"><div align="center">Component Type </div></td>
  <td class="t-t-Borrow"><div align="center">Cells Number </div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
STATION_ID=rs("STATION_ID")
while not rs.eof and i<=rs.pagesize 
if STATION_ID<>rs("STATION_ID") then
%>
<tr><td height="5" colspan="21" class="today">&nbsp;</td>
</tr>
<%end if%>
<tr>
  <td width="40" height="20"><div align="center"><% =(cint(session("strpagenum"))-1)*recordsize+i%></div></td>
   
    <td width="30" height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditAction.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td width="46" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if(confirm('Are you sure to delete this Action?')){window.open('DeleteAction.asp?id=<%=rs("NID")%>&actionname=<%=rs("ACTION_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	 <td width="144" height="20"><div align="center"><%= rs("NID") %></div></td>
	<td width="144"><div align="center" class="red">
	  <%if rs("STATUS")="1" then%>
	    <span style="cursor:hand" onClick="javascript:location.href='DisableAction.asp?id=<%=rs("NID")%>&actionname=<%=rs("ACTION_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Action"><img src="/Images/Enabled.gif"></span>
	    <%else%>
	    <span style="cursor:hand" onClick="javascript:location.href='EnableAction.asp?id=<%=rs("NID")%>&actionname=<%=rs("ACTION_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Action"><img src="/Images/Disabled.gif"></span>
        <%end if%>
    </div></td>
	
    <td width="144" height="20"><div align="center"><a href="/Admin/Station/EditStation.asp?id=<%=rs("STATION_ID")%>&path=<%=path%>&query=<%=query%>"><%= rs("STATION_NAME") %>&nbsp;(<%=rs("STATION_CHINESE_NAME")%>)</a></div></td>
    <td width="144" height="20"><div align="center"><%= rs("ACTION_NAME") %></div></td>
	<td width="223"><div align="center"><%= rs("ACTION_CHINESE_NAME") %></div></td>
	<td width="223"><div align="center"><%= rs("FACTORY_NAME") %></div></td>
	<%select case rs("ACTION_PURPOSE")
	case "0"
	action_purpose="Other"
	case "1"
	action_purpose="Machine Code"
	case "2"
	action_purpose="Material Part Number"
	case "3"
	action_purpose="Material Lot Number"
	case "4"
	action_purpose="Material Quanity"
	case else
	action_purpose=""
	end select%>
    <td width="223"><div align="center"><%= action_purpose %>&nbsp;</div></td>
    <td width="223"><div align="center">
        <%if rs("STATION_POSITION")="0" then%>
        B
		<%elseif rs("STATION_POSITION")="1" then%>
		A
        <%else%>
        &nbsp;
        <%end if%>
    </div></td>
    <td width="223"><div align="center"><%if rs("NULL_ALLOW")="0" then%><img src="/Images/Delete.gif"><%else%><img src="/Images/Finished.gif"><%end if%>&nbsp;</div></td>
	<td width="223"><div align="center"><%if rs("APPEND_ALLOW")="0" then%>N<%else%>Y<%end if%>&nbsp;</div></td>
    <td width="223"><div align="center"><%if rs("WITH_LOT")="1" then%>Y<%else%>N<%end if%>&nbsp;</div></td>
    <td width="223"><div align="center"><%= rs("VALID_MACHINE") %>&nbsp;</div></td>
    <td width="223"><%=getAction(true,"TEXT",""," where A.NID='"&rs("RELATED_ACTION_ID")&"'","","","")%>&nbsp;</td>
    <td width="223"><div align="center"><%= rs("MAX_QUANTITY") %>&nbsp;</div></td>
    <td width="223"><div align="center"><%= rs("MIN_QUANTITY") %>&nbsp;</div></td>
    <td width="223"><div align="center"><%= rs("ACTION_TYPE") %>&nbsp;</div></td>
    <td width="223"><div align="center"><%= rs("ELEMENT_TYPE") %></div></td>
    <td width="223"><div align="center"><%= rs("ELEMENT_NUMBER") %></div></td>
</tr>
<%
i=i+1
STATION_ID=rs("STATION_ID")
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="21"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->