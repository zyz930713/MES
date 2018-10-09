<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
server.ScriptTimeout=200%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
create_time=now()
rnd_key=gen_key(10)
thiserror=""

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.form("factory")
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
if fromdate="" then
fromdate=dateadd("d",-7,date())
end if
if todate="" then
todate=date()
end if
if fromhour="" then
fromhour="0:00:00"
end if
if tohour="" then
tohour="0:00:00"
end if
fromtime=fromdate&" "&fromhour
totime=todate&" "&tohour
pagename="/Reports/Process/FamilyScrap/GenerateFamilyScrap.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="FAMILY_SCRAP"
cmd.CommandType=4
'response.Write(factory&"#"&fromtime&"#"&totime&"#"&session("code")&"#"&rnd_key)
'response.End()
cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 10, factory)
cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, fromtime)
cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, totime)
cmd.Parameters.Append cmd.CreateParameter("creator_code", adVarChar, adParamInput, 4, session("code"))
cmd.Parameters.Append cmd.CreateParameter("rnd_key", adVarChar, adParamInput, 10, rnd_key)
cmd.execute
set cmd=nothing
if err.number=0 then
word="Fail to create report.\n创建报表失败！"
end if
SQL="select TARGET_YIELD from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_yield=rs("TARGET_YIELD")
end if
rs.close
FactoryRight "S."
SQL="select * from FAMILY_SCRAP_DETAIL_TEMP where CREATOR_CODE='"&session("code")&"' and RND_KEY='"&rnd_key&"' order by FAMILY_NAME"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/Process/FamilyScrap/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy">Browse Family Scrap (from <%=fromtime%> to <%=totime%>) </td>
  </tr>
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy">User:
    <% =session("User") %></td>
  </tr>
  <tr>
    <td height="20" colspan="8">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Family Name </div></td>
    <td class="t-t-Borrow"><div align="center">Assembly Input </div></td>
    <td class="t-t-Borrow"><div align="center"> Scrap Quantity </div></td>
    <td class="t-t-Borrow"><div align="center">Scrap Amount </div></td>
    <td class="t-t-Borrow"><div align="center"> Scrap  Percetage </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Target Percetage </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Target Amount </div></td>
  </tr>
  <%
i=1
if not rs.eof then
%>
<form action="/Reports/Process/FamilyScrap/FamilyScrapDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_input=0
total_scrap_quantity=0
total_scrap_amount=0
while not rs.eof
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.jobs.value='<%=rs("JOBS")%>';document.all.seriesgroup.value='<%=rs("FAMILY_NAME")%>';document.formDetail.submit()"><%=rs("FAMILY_NAME")%></span></td>
    <td height="20"><div align="center"><%=rs("INPUT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("SCRAP_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("SCRAP_AMOUNT")%></div></td>
    <td height="20" class="<%=GetScrapColor(rs("SCRAP_PERCENTAGE"),rs("FAMILY_TARGET_PERCENTAGE"))%>"><div align="center"><%=formatpercent(csng(rs("SCRAP_PERCENTAGE")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("FAMILY_TARGET_PERCENTAGE")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=rs("FAMILY_TARGET_AMOUNT")%></div></td>
    </tr>
  <%
	total_input=total_input+csng(rs("INPUT_QUANTITY"))
	total_scrap_quantity=total_scrap_quantity+csng(rs("SCRAP_QUANTITY"))
	total_scrap_amount=total_scrap_amount+csng(rs("SCRAP_AMOUNT"))
rs.movenext
i=i+1
wend
%> 
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td height="20"><div align="center"><%=total_input%></div></td>
    <td height="20"><div align="center"><%=total_scrap_quantity%></div></td>
    <td height="20"><div align="center"><%=total_scrap_amount%></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
<input name="jobs" type="hidden" value="">
<input name="factory_id" type="hidden" value="<%=factory%>">
<input name="from_time" type="hidden" value="<%=fromtime%>">
<input name="to_time" type="hidden" value="<%=totime%>">
<input name="seriesgroup" type="hidden" value="">
</form>
  <form name="form1" method="post" action="/Reports/Process/FamilyScrap/SaveFamilyScrap.asp" onSubmit="return formcheckSave()">
    <tr>
      <td height="20" colspan="8">Generating time:
        <% =formatdate(create_time,application("longdateformat"))%>
      &nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="8">Chart Report:
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
      <td height="20" colspan="8">Report name:        
<input name="finalfamily_name" type="text" id="finalfamily_name">
        <input name="factory_target_yield" type="hidden" id="factory_target_yield" value="<%=factory_target_yield%>">
        <input name="factory_target_amount" type="hidden" id="factory_target_amount" value="<%=factory_target_amount%>">
        <input name="from_time" type="hidden" id="from_time" value="<%=fromtime%>">
        <input name="to_time" type="hidden" id="to_time" value="<%=totime%>">
		<input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
        <input name="factory" type="hidden" id="factory" value="<%=factory%>">
		<input name="total_input" type="hidden" id="total_input" value="<%=total_input%>">
		<input name="total_scrap_quantity" type="hidden" id="total_scrap_quantity" value="<%=total_scrap_quantity%>">
		<input name="total_scrap_amount" type="hidden" id="total_scrap_amount" value="<%=total_scrap_amount%>">
		<input name="total_scrap_perecentage" type="hidden" id="total_scrap_perecentage" value="<%=total_scrap_perecentage%>">
		<input name="Save" type="submit" id="Save" value="Save This Report">      </td>
    </tr>
  </form>
  <%
else
%>
  <tr>
    <td height="20" colspan="8"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->