<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Group/GroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
id=request.QueryString("id")
group_name=trim(request("group_name"))
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Admin/Group/ValidareGroup.asp.asp"
SQL="select GROUP_MEMBERS from SYSTEM_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
members=rs("GROUP_MEMBERS")
end if
rs.close
action_id=""
action_name=""
action_machine=""
astation_id=""
astation_name=""
SQL="select 1,GA.ACTION_ID,GA.MACHINE,A.ACTION_NAME,A.STATION_ID,S.STATION_NAME from GROUP_ACTION_VALUE GA inner join ACTION A on GA.ACTION_ID=A.NID inner join STATION S on A.STATION_ID=S.NID where GA.GROUP_ID='"&id&"' and MACHINE is not null"
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
	action_id=action_id&rs("ACTION_ID")&","
	action_name=action_name&rs("ACTION_NAME")&","
	action_machine=action_machine&rs("MACHINE")&";"
	astation_id=astation_id&rs("STATION_ID")&","
	astation_name=astation_name&rs("STATION_NAME")&","
	rs.movenext
	wend
end if
rs.close

action_id=left(action_id,len(action_id)-1)
action_name=left(action_name,len(action_name)-1)
action_machine=left(action_machine,len(action_machine)-1)
astation_id=left(astation_id,len(astation_id)-1)
astation_name=left(astation_name,len(astation_name)-1)

aaction_id=split(action_id,",")
aaction_name=split(action_name,",")
aaction_machine=split(action_machine,";")
aastation_id=split(astation_id,",")
aastation_name=split(astation_name,",")

SQL="select 1,J.JOB_NUMBER from JOB J where J.JOB_NUMBER in ('"&replace(members,",","','")&"') order by J.JOB_NUMBER"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=ubound(aaction_id)+3%>" class="t-c-greenCopy">Result of validate group -- <%=group_name%></td>
</tr>
<tr>
  <td height="20" colspan="<%=ubound(aaction_id)+3%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=ubound(aaction_id)+3%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number </div></td>
  <%for j=0 to ubound(aaction_id)%>
  <td class="t-t-Borrow"><div align="center"><span title="<% =aaction_machine(j) %>"><%=aaction_name(j)%> in <%=aastation_name(j)%></span></div></td>
  <%next%>
  </tr>
<%
i=1
ispass=true
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><%=i%></div></td>
    <td><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("JOB_NUMBER") %></a></div></td>
	<%for j=0 to ubound(aaction_id)%>
    <td><div align="center">
	<%this_action_value=getJobActionValue(rs("JOB_NUMBER"),aastation_id(j),aaction_id(j))
	if instr(aaction_machine(j),this_action_value)=0 then
	ispass=false
	%><span class="red"><%=this_action_value%></span><%else%><%=this_action_value%><%end if%></div></td>
	<%next%>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
  <tr>
    <td height="20" colspan="<%=ubound(aaction_id)+3%>">Result: <%if ispass=true then%><span class="green">Pass Invalidation</span><%else%><span class="red">Fail</span><%end if%></td>
  </tr>
<%else%>
  <tr>
    <td height="20" colspan="3"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->