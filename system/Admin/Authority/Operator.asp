<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
usercode=trim(request("usercode"))
username=trim(request("username"))
ordername=request("ordername")
ordertype=request("ordertype")
force=request("force")
forcefactory=request("forcefactory")
factory=request.QueryString("factory")
line=request.QueryString("line")
if ordername="" and ordertype="" then
order=" order by O.CODE"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if usercode<>"" then
where=where&" and UPPER(O.CODE)='"&ucase(usercode)&"'"
end if
if username<>"" then
where=where&" and O.OPERATOR_NAME like '%"&username&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and O.FACTORY_ID is null"
else
where=where&" and O.FACTORY_ID='"&factory&"'"
end if
if line="" or line="all" then
where=where&""
elseif line="null" then
where=where&" and O.LINE_ID is null"
else
where=where&" and O.LINE_ID='"&line&"'"
end if
pagepara="&usercode="&usercode&"&username="&username&"&factory="&factory&"&line="&line&"&force="&force&"&forcefactory="&forcefactory
pagename="/Admin/Authority/Operator.asp"
if force="1" then
SQL="Select O.*,F.FACTORY_NAME,L.LINE_NAME from OPERATORS O left join FACTORY F on O.FACTORY_ID=F.NID left join LINE L on O.LINE_ID=L.NID where UPPER(O.CODE)='"&ucase(usercode)&"'"
else
FactoryRight "O."
SQL="Select O.*,F.FACTORY_NAME,L.LINE_NAME from OPERATORS O left join FACTORY F on O.FACTORY_ID=F.NID left join LINE L on O.LINE_ID=L.NID where O.CODE is not null"&where&factorywhereoutsideand&order
end if
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Admin/Authority/Lan_Operator.asp" -->
<script language="javascript">
function forcecheck()
{
	with(document.form1)
	{
		if(force.checked)
		{
		forcefactory.disabled=false;
		}
		else
		{
		forcefactory.disabled=true;
		forcefactory.options[0].selected;
		}
	}
}
function formcheck()
{
	with(document.form1)
	{
		if(force.checked&&forcefactory.selectedIndex==0)
		{
		alert("Please select factory.\n请为借调人员选择工厂！");
		return false;
		}
		if(force.checked&&usercode.value=="")
		{
		alert("Please key in code.\n请输入借调人员工号！");
		return false;
		}
	}
}
</script>
</head>

<body onLoad="language();language_page();forcecheck()">
<form action="/Admin/Authority/Operator.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span class="style1"><span id="inner_SearchCode"></span></td>
    <td height="20"><input name="usercode" type="text" id="usercode" value="<%=usercode%>">
      <input name="force" type="checkbox" id="force" onClick="forcecheck()" value="1" <%if force="1" then%>checked<%end if%>>
      <span id="inner_SearchForcedCode"></span>
      <select name="forcefactory" disabled id="forcefactory">
        <option></option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",forcefactory,factorywhereinside,"","") %>
      </select></td>
    <td><span id="inner_SearchEnglishName"></span></td>
    <td><input name="username" type="text" id="username" value="<%=username%>"></td>
    <td><span id="inner_SearchChineseName"></span></td>
    <td><input name="userchinesename" type="text" id="userchinesename" value="<%=userchinesename%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">
      </td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy">    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><input name="byStation" type="button" id="byStation" onClick="javascript:location.href='StationOperator.asp'" value="Browse by Station">
          <input name="byPart" type="button" id="byPart" onClick="javascript:location.href='PartOperator.asp'" value="Browse by Part"></td>
        <td><div align="right"><%if admin=true then%><a href="/Admin/Authority/NewOperator.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_Add"></span></a><%end if%></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td colspan="2" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_Locked"></span></div></td>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_Practised"></span></div></td>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_Period"></span></div></td>
  <%end if%>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=O.CODE&ordertype=asc<%=pagepara%>'"><span id="inner_Code"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=O.CODE&ordertype=desc<%=pagepara%>'"></div></td>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=O.OPERATOR_NAME&ordertype=asc<%=pagepara%>'"><span id="inner_EnglishName"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=O.OPERATOR_NAME&ordertype=desc<%=pagepara%>'"></div></td>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_ChineseName"></span></div></td>
  <td nowrap class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value+'&line=<%=line%>'">
      <option value=""></option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td nowrap class="t-t-Borrow"><select name="Line" id="Line" onChange="location.href='<%=pagename%>?factory=<%=factory%>&line='+this.options[this.selectedIndex].value">
    <option value=""></option>
    <option value="all" <%if line="all" then%>selected<%end if%>>All</option>
    <option value="null" <%if line="null" then%>selected<%end if%>>None</option>
    <%FactoryRight "L."%>
    <%= getLine("OPTION",line,factorywhereoutside,"","") %>
	 </select>
 </td>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
  <td nowrap class="t-t-Borrow"><div align="center"><span id="inner_Parts"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr <%if rs("PRACTISED")="1" then%>class="t-c-practised"<%end if%>>
  <td><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditOperator.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
	<td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('DeleteOperator.asp?id=<%=rs("NID")%>&operatorname=<%=rs("OPERATOR_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	<td><div align="center">
        <%if rs("LOCKED")="1" then%>
	  <img src="/Images/OLocked.gif" alt="Locked" width="12" height="15">
	  <%else%>
	  <img src="/Images/OUnLocked.gif" alt="UNlocked" width="12" height="15">
	  <%end if%>
    </div></td>
	<td><div align="center">
	  <div align="center">
	    <%if rs("PRACTISED")="1" then%>
	    <img src="/Images/Yes.gif" width="16" height="16" align="absbottom">
	    <%end if%>&nbsp;      </div>
	</div></td>
	<td><div align="center"><%if rs("PRACTISED")="1" then%><%= rs("PRACTISE_START_TIME") %> - <%= rs("PRACTISE_END_TIME") %><%end if%>
	  &nbsp;</div></td>
	<%end if%>
    <td height="20"><div align="center" class="red"><%if admin=true then%><span style="cursor:hand" onClick="javascript:window.open('EditAuthority.asp?code=<%=rs("code")%>&force=<%=force%>&forcefactory=<%=forcefactory%>&path=<%=path%>&query=<%=query%>','main')"><%= rs("CODE") %></span><%else%><%= rs("CODE") %><%end if%></div></td>
    <td height="20"><div align="center"><%= rs("OPERATOR_NAME") %></div></td>
    <td><div align="center"><%= rs("OPERATOR_CHINESE_NAME") %></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("LINE_NAME") %>&nbsp;</div></td>
    <td><div align="left">
	<%if rs("AUTHORIZED_STATIONS_ID")<>"" then%>
	<%= getStation_New(true,"TEXT",""," where S.NID in ('"&replace(rs("AUTHORIZED_STATIONS_ID"),",","','")&"')","",rs("AUTHORIZED_STATIONS_ID")," ; ") %>
	<%end if%>
	&nbsp;</div></td>
    <td><%if rs("AUTHORIZED_PARTS_ID")<>"" then%>
      <%= getPart(true,"TEXT",""," where P.NID in ('"&replace(rs("AUTHORIZED_PARTS_ID"),",","','")&"')","",rs("AUTHORIZED_PARTS_ID")," ; ","","","","","","") %>
      <%end if%>
&nbsp;</td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
<tr>
    <td height="20" colspan="13"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
