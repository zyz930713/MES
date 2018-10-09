<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Action/ActionCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Admin/Action/Action_New.asp"
actionname=request("actionname")
chineseactionname=request("chineseactionname")
thisstatus=request.QueryString("status")
station=request("station")
factory=request.QueryString("factory")
purpose=request.QueryString("purpose")
actioncode=request("actioncode")

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
pagepara="&status="&thisstatus&"&station="&station&"&factory="&factory&"&purpose="&purpose
FactoryRight "A."
SQL="select A.NID,A.NID,A.ACTION_NAME,A.ACTION_CODE,A.ACTION_CHINESE_NAME,A.ACTION_PURPOSE,A.STATION_POSITION,A.APPEND_ALLOW,A.NULL_ALLOW,A.ACTION_TYPE,A.ELEMENT_TYPE,A.ELEMENT_NUMBER,A.WITH_LOT,F.FACTORY_NAME from ACTION_New A  left join Factory F on A.FACTORY_ID=F.NID where IS_DELETE=0 "&where&factorywhereoutsideand&" Order by A.ACTION_CODE"
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
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Admin/Action/Action_New.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_Name"></span></td>
    <td height="20"><input name="actionname" type="text" id="actionname" value="<%=actionname%>"></td>
	<td height="20"><span id="inner_ActionCode"></span></td>
    <td height="20"><input name="actioncode" type="text" id="actioncode" value="<%=actioncode%>"></td>
	
    <td><span id="inner_CHName"></span></td>
    <td><input name="chineseactionname" type="text" id="chineseactionname" value="<%=chineseactionname%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="20" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="20" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%"><span id="inner_User"></span>: <% =session("User") %></td>
        <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Action/AddAction_New.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a><%end if%></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="20"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td> 
  <%end if%> 
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Name"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_ActionCode"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_CHName"></span></div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?&status=<%=thisstatus%>&factory='+this.options[this.selectedIndex].value+'&purpose=<%=purpose%>'">
      <option value="">Factory</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="purpose" id="purpose" onChange="location.href='<%=pagename%>?status=<%=thisstatus%>&factory=<%=factory%>&purpose='+this.options[this.selectedIndex].value">
      <option value=""></option>
	  <%for i=1 to Ubound(ActionPurpose)
	  		response.Write("<option value='"&i&"'")
			if purpose = cstr(i) then
				response.Write("selected")
			end if
			response.Write(" >"&ActionPurpose(i)&"</option>")
	  	next
	  %>      
      <option value="0" <%if purpose="0" then%>selected<%end if%>>Other</option>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Position"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Null"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Append"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_HasLotProperty"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ActionType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ComponentType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_CellNo"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td width="40" height="20"><div align="center"><% =(cint(session("strpagenum"))-1)*recordsize+i%></div></td>
  <%if admin=true then%>
    <td width="30" height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditAction_New.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td width="46" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if(confirm('Are you sure to delete this Action?')){window.open('DeleteAction_New.asp?id=<%=rs("NID")%>&actionname=<%=rs("ACTION_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
    <%end if%>
    <td width="144" height="20"><div align="center"><%= rs("ACTION_NAME") %></div></td>
	<td width="144" height="20"><div align="center"><%= rs("ACTION_CODE") %></div></td>
	<td width="223"><div align="center"><%= rs("ACTION_CHINESE_NAME") %></div></td>
	<td width="223"><div align="center"><%= rs("FACTORY_NAME") %>&nbsp;</div></td>	
    <td width="223"><div align="center">
	<%if cint(rs("ACTION_PURPOSE"))>=0 and cint(rs("ACTION_PURPOSE"))<= Ubound(ActionPurpose) then 
		response.Write(ActionPurpose(rs("ACTION_PURPOSE")))
	  end if%>&nbsp;</div></td>
    <td width="223"><div align="center">
        <%if rs("STATION_POSITION")="0" then%>
        Before
		<%elseif rs("STATION_POSITION")="1" then%>
		After
        <%else%>
        &nbsp;
        <%end if%>
    </div></td>
    <td width="223"><div align="center"><%if rs("NULL_ALLOW")="0" then%><img src="/Images/Delete.gif"><%else%>
      <img src="/Images/Finished.gif">
      <%end if%>
      &nbsp;</div></td>	
	<td width="223"><div align="center"><%if rs("APPEND_ALLOW")="0" then%>N<%else%>Y<%end if%>&nbsp;</div></td>
    <td width="223"><div align="center"><%if rs("WITH_LOT")="1" then%>Y<%else%>N&nbsp;<%end if%></div></td>      
    <td width="223"><div align="center"><%= rs("ACTION_TYPE") %>&nbsp;</div></td>
    <td width="223"><div align="center"><%= rs("ELEMENT_TYPE") %></div></td>
    <td width="223"><% if rs("ELEMENT_NUMBER") <> "" then %><div align="center"><%= rs("ELEMENT_NUMBER") %></div><%else%>&nbsp;<%end if%></td>
</tr>
<%
i=i+1
rs.movenext
wend 
rs.close
else
%>
  <tr>
    <td height="20" colspan="20"><div align="center">No Records </div></td>
  </tr>
</table>
<% 
end if%>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->