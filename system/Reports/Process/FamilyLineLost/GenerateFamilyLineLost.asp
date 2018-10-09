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
pagename="/Reports/Process/FamilyLineLost/GenerateFamilyLineLost.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="FAMILY_LINELOST"
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
word="Fail to create schedule.\n创建计划任务失败！"
end if
SQL="select TARGET_LOST_QUANTITY,TARGET_LOST_AMOUNT from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_quantity=rs("TARGET_LOST_QUANTITY")
factory_target_amount=rs("TARGET_LOST_AMOUNT")
end if
rs.close
FactoryRight "S."
SQL="select * from FAMILY_LINELOST_DETAIL_TEMP where CREATOR_CODE='"&session("code")&"' and RND_KEY='"&rnd_key&"' order by FAMILY_NAME"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/Process/FamilyLineLost/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy">Browse Family Line Lost (from <%=fromtime%> to <%=totime%>) </td>
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
    <td class="t-t-Borrow"><div align="center"> Lost Quantity </div></td>
    <td class="t-t-Borrow"><div align="center">Lost Amount </div></td>
    <td class="t-t-Borrow"><div align="center"> Lost  Percetage </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Target Quantity</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Target Amount </div></td>
  </tr>
  <%
i=1
if not rs.eof then
%>
<form action="/Reports/Process/FamilyLineLost/FamilyLineLostDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_input=0
total_lost_quantity=0
total_lost_amount=0
while not rs.eof
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.included_system_items.value='<%=rs("INCLUDED_SYSTEM_ITEMS")%>';document.all.seriesgroup.value='<%=rs("FAMILY_NAME")%>';document.formDetail.submit()"><%=rs("FAMILY_NAME")%></span></td>
    <td height="20"><div align="center"><%=rs("INPUT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("LINELOST_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("LINELOST_AMOUNT")%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("LINELOST_PERCENTAGE")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=rs("FAMILY_TARGET_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("FAMILY_TARGET_AMOUNT")%></div></td>
    </tr>
  <%
	total_input=total_input+csng(rs("INPUT_QUANTITY"))
	total_lost_quantity=total_lost_quantity+csng(rs("LINELOST_QUANTITY"))
	total_lost_amount=total_lost_amount+csng(rs("LINELOST_AMOUNT"))
rs.movenext
i=i+1
wend
%> 
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td height="20"><div align="center"><%=total_input%></div></td>
    <td height="20"><div align="center"><%=total_lost_amount%></div></td>
    <td height="20"><div align="center"><%=total_lost_amount%></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    </tr>
<input name="included_system_items" type="hidden" value="">
<input name="factory_id" type="hidden" value="<%=factory%>">
<input name="from_time" type="hidden" value="<%=fromtime%>">
<input name="to_time" type="hidden" value="<%=totime%>">
<input name="seriesgroup" type="hidden" value="">
</form>
  <form name="form1" method="post" action="/Reports/Process/FamilyLineLost/SaveFamilyLineLost.asp" onSubmit="return formcheckSave()">
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
        <input name="factory_target_quantity" type="hidden" id="factory_target_quantity" value="<%=factory_target_quantity%>">
        <input name="factory_target_amount" type="hidden" id="factory_target_amount" value="<%=factory_target_amount%>">
        <input name="from_time" type="hidden" id="from_time" value="<%=fromtime%>">
        <input name="to_time" type="hidden" id="to_time" value="<%=totime%>">
		<input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
        <input name="factory" type="hidden" id="factory" value="<%=factory%>">
		<input name="total_input" type="hidden" id="total_input" value="<%=total_input%>">
		<input name="total_lost_quantity" type="hidden" id="total_lost_quantity" value="<%=total_lost_quantity%>">
		<input name="total_lost_amount" type="hidden" id="total_lost_amount" value="<%=total_lost_amount%>">
		<input name="total_lost_perecentage" type="hidden" id="total_lost_perecentage" value="<%=total_lost_perecentage%>">
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