<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
server.ScriptTimeout=500%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/OOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
create_time=now()
rnd_key=gen_key(10)
thiserror=""

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.form("factory")
section=request.form("section")
fromdate=request.form("jfromdate")
todate=request.form("jtodate")
fromhour=request.form("jfromhour")
tohour=request.form("jtohour")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by L.LINE_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" then
fromdate=dateadd("d",-1,date())
end if
if todate="" then
todate=date()
end if
if fromhour="" then
fromhour="8:30:00"
end if
if tohour="" then
tohour="8:30:00"
end if
fromtime=fromdate&" "&fromhour
totime=todate&" "&tohour
pagename="/Reports/Process/Output/GenerateOutput.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="FAMILY_OUTPUT"
cmd.CommandType=4
'response.Write(factory&"#"&section&"#"&fromtime&"#"&totime&"#"&session("code")&"#"&rnd_key)
'response.End()
cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 10, factory)
cmd.Parameters.Append cmd.CreateParameter("v_section_id", adVarChar, adParamInput, 10, section)
cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, fromtime)
cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, totime)
cmd.Parameters.Append cmd.CreateParameter("creator_code", adVarChar, adParamInput, 4, session("code"))
cmd.Parameters.Append cmd.CreateParameter("rnd_key", adVarChar, adParamInput, 10, rnd_key)
cmd.execute
set cmd=nothing
if err.number=0 then
word="Fail to create report.\n创建报表失败！"
end if
family_name=getSeriesGroup("TEXT",null," where S.FACTORY_ID='"&factory&"'","",",")
station_name=getStation(false,"ENGLISH_TEXT",null," where S.OUTPUT_REPORT_COLUMN=1 and S.SECTION_ID='"&section&"' and S.FACTORY_ID='"&factory&"'"," order by S.OUTPUT_SEQUENCY",null,",")
if family_name<>"" then
family_name=left(family_name,len(family_name)-1)
a_family_name=split(family_name,",")
end if
if station_name<>"" then
a_station_name=split(station_name,",")
end if
dim total_station_quantity()
redim total_station_quantity(ubound(a_station_name))
Tcount=ubound(a_station_name)+3
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
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Family Scrap (from <%=fromtime%> to <%=totime%>) </td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Famlily Name</div></td>
  <%for i=0 to ubound(a_station_name)%>
  <td class="t-t-Borrow"><div align="center"><%=a_station_name(i)%></div></td>
  <%next%>
  </tr>
<%for i=0 to ubound(a_family_name)%>
<tr>
  <td height="20"><div align="center"><%=i+1%></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('OutputLineDetail.asp?line_name=<%=a_family_name(i)%>&fromtime=<%=fromtime%>&totime=<%=totime%>&output_name=<%=output_name%>&output_report_time=<%=output_report_time%>&section_id=<%=section%>')"><%=a_family_name(i)%></span></div></td>
    <%for j=0 to ubound(a_station_name)
		SQL="select OUTPUT from FAMILY_OUTPUT_DETAIL_TEMP where FAMILY_NAME='"&a_family_name(i)&"' and STATION_NAME='"&a_station_name(j)&"'"
		'response.Write(SQL)
		rs.open SQL,conn,1,3
		if not rs.eof then
		output=rs("OUTPUT")
		else
		output=0
		end if
		rs.close
		total_station_quantity(j)=total_station_quantity(j)+csng(output)%>
	  <div align="center"><td><div align="center"><%=output%>&nbsp;K</div></td></div>
	<%next%>
  </tr>
<%next%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <%for j=0 to ubound(total_station_quantity)%>
  <td><div align="center">
      <% =total_station_quantity(j)%>
&nbsp;</div></td>
  <%next%>
</tr>
      <form name="form1" method="post" action="/Reports/Process/FamilyOutput/SaveFamilyOutput.asp">
  
  <tr>
    <td height="20" colspan="<%=Tcount%>">Generating time: <% =formatdate(create_time,application("longdateformat"))%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>">Chart Report:
      <input name="isWW" type="checkbox" id="isWW" value="1">
Save this report as WW
<input name="wwnumber" type="text" id="wwnumber" size="2" onChange="weeknumbercheck()">
of
<input name="yearnumber" type="text" id="yearnumber" value="<%=year(date())%>" size="4" onChange="yearnumbercheck()">
year
<input name="monthnumber" type="text" id="monthnumber" value="<%=month(date())%>" size="2" onChange="monthnumbercheck()">
month </td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>">
        Report name: 
        <input name="output_name" type="text" id="output_name">
        <input name="fromtime" type="hidden" id="fromtime" value="<%=fromtime%>">
        <input name="totime" type="hidden" id="totime" value="<%=totime%>">
        <input name="factory" type="hidden" id="factory" value="<%=factory%>">
        <input name="section" type="hidden" id="section" value="<%=section%>">
        <input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
        <input name="Save" type="submit" id="Save" value="Save This Report">    </td>
  </tr> 
  </form>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/OOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->