<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
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
pagename="/Reports/FinalYield/WeeklyFinanceYiled/GenerateWeeklyFinanceYield.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="RUN_WEEKLY_FINANCEYIELD"
cmd.CommandType=4 
'response.Write(factory&"*"&fromtime&"*"&totime&"*"&session("code")&"*"&rnd_key)
'response.End()
cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 10, factory)
cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, fromtime)
cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, totime)
cmd.Parameters.Append cmd.CreateParameter("creator_code", adVarChar, adParamInput, 4, session("code"))
cmd.Parameters.Append cmd.CreateParameter("rnd_key", adVarChar, adParamInput, 10, rnd_key)
cmd.execute
set cmd=nothing
if err.number=0 then
word="Fail to create report.\n生成报告失败！"
end if
SQL="select FINANCE_TARGET_YIELD,FINANCE_PLAN_TARGET_YIELD from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_finance_ield=rs("FINANCE_TARGET_YIELD")
factory_plan_target_finance_ield=rs("FINANCE_PLAN_TARGET_YIELD")
end if
rs.close
FactoryRight "S."
SQL="select * from BAR_REPORT.WEEKLY_FINANCE_YIELD_TEMP where CREATOR_CODE='"&session("code")&"' and RND_KEY='"&rnd_key&"' order by FAMILY_NAME"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/FinalYield/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy">Browse Weekly Finance Yield (from <%=fromtime%> to <%=totime%>) </td>
  </tr>
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy">User:
    <% =session("User") %></td>
  </tr>
  <tr>
    <td height="20" colspan="9">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Family Name </div></td>
    <td class="t-t-Borrow"><div align="center"> Input Value</div></td>
    <td class="t-t-Borrow"><div align="center"> Output Value </div></td>
    <td class="t-t-Borrow"><div align="center"> Scrap Value </div></td>
    <td class="t-t-Borrow"><div align="center">Yield </div></td>
    <td class="t-b-newyear"><div align="center"> Target Yield </div></td>
    <td class="t-t-Borrow"><div align="center">Plan Yield </div></td>
    <td class="t-t-Borrow"><div align="center">Delta</div></td>
  </tr>
  <%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalFamilyYield/FinalFamilyDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_input_amount=0
total_output_amount=0
total_scrap_amount=0
while not rs.eof
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><div align="center"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.family_name.value='<%=rs("FAMILY_NAME")%>';document.formDetail.submit()"><%=rs("FAMILY_NAME")%></span></div></td>
    <td height="20"><div align="center"><%=rs("INPUT_AMOUNT")%></div></td>
    <td height="20"><div align="center"><%=rs("OUTPUT_AMOUNT")%></div></td>
    <td><div align="center"><%=rs("SCRAP_AMOUNT")%></div></td>
    <td height="20" class="<%=getYieldColor(rs("AMOUNT_YIELD"),rs("TARGET"))%>"><div align="center"><%=formatpercent(csng(rs("AMOUNT_YIELD")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("PLAN_TARGET"))/100,2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(csng(rs("TARGET"))/100,2,-1)%></div></td>
    <td><div align="center">
    <%if csng(rs("AMOUNT_YIELD"))<>0 then
	delta=csng(rs("AMOUNT_YIELD"))-csng(rs("TARGET"))/100
	else
	delta=0
	end if%>
    <%=formatpercent(delta,2,-1)%></div></td>
  </tr>
  <%
	total_input_amount=total_input_amount+csng(rs("INPUT_AMOUNT"))
	total_output_amount=total_output_amount+csng(rs("OUTPUT_AMOUNT"))
	total_scrap_amount=total_scrap_amount+csng(rs("SCRAP_AMOUNT"))
rs.movenext
i=i+1
wend
if total_output_amount<>0 then
total_yield=total_output_amount/total_input_amount
else
total_yield=0
end if
%> 
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td height="20"><div align="center"><%=total_input_amount%></div></td>
    <td height="20"><div align="center"><%=total_output_amount%></div></td>
    <td><div align="center"><%=total_scrap_amount%></div></td>
    <td height="20" class="<%=getYieldColor(total_yield,factory_target_yield)%>"><div align="center"><%=formatpercent(total_yield,2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(csng(factory_target_yield)/100,2,-1)%>
      </div>
    </div></td>
    <td><div align="center"><%=formatpercent(csng(factory_plan_target_yield)/100,2,-1)%></div></td>
    <td><div align="center">
      <%if csng(total_final_yield)<>0 then
	delta=csng(total_final_yield)-csng(factory_target_yield)/100
	else
	delta=0
	end if%>
      <%=formatpercent(delta,2,-1)%></div></td>
  </tr>
<input name="family_name" type="hidden" value="">
</form>
  <tr>
    <td height="20" colspan="9"><table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#333333" bordercolordark="#FFFFFF">
      <tr height="21">
        <td rowspan="3">Final Yield </td>
        <td width="50"><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Green">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td width="201">Means cum yield meet target</td>
      </tr>
      <tr height="21">
        <td><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Yellow">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means delta is between 0% to 0.5%</td>
      </tr>
      <tr height="21">
        <td><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Red">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means delta is bigger than 0.5%</td>
      </tr>
    </table></td>
  </tr>
  <form name="form1" method="post" action="/Reports/FinalYield/WeeklyFinanceYield/SaveWeeklyFinanceYield.asp" onSubmit="return formcheckSave()">
    <tr>
      <td height="20" colspan="9">Generating time:
        <% =formatdate(create_time,application("longdateformat"))%>
      &nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="9">Chart Report:
        <input name="isWW" type="checkbox" id="isWW" value="1">
Save this report as WW
  <input name="wwnumber" type="text" id="wwnumber" size="2" onChange="weeknumbercheck(this)">
of  
	<input name="yearnumber" type="text" id="yearnumber" value="<%=year(date())%>" size="4" onChange="yearnumbercheck(this)">
year
<input name="monthnumber" type="text" id="monthnumber" value="<%=month(date())%>" size="2" onChange="monthnumbercheck(this)">
month </td>
    </tr>
    <tr>
      <td height="20" colspan="9">Report name:        
<input name="finance_yield_name" type="text" id="finance_yield_name">
        <input name="factory_target_yield" type="hidden" id="factory_target_yield" value="<%=factory_target_yield%>">
        <input name="factory_plan_target_yield" type="hidden" id="factory_plan_target_yield" value="<%=factory_plan_target_yield%>">
        <input name="from_time" type="hidden" id="from_time" value="<%=fromtime%>">
        <input name="to_time" type="hidden" id="to_time" value="<%=totime%>">
		<input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
        <input name="factory" type="hidden" id="factory" value="<%=factory%>">
		<input name="total_input_amount" type="hidden" id="total_input_amount" value="<%=total_input_amount%>">
		<input name="total_output_amount" type="hidden" id="total_output_amount" value="<%=total_output_amount%>">
		<input name="total_scrap_amount" type="hidden" id="total_scrap_amount" value="<%=total_scrap_amount%>">
		<input name="total_yield" type="hidden" id="total_yield" value="<%=total_yield%>">
      <input name="Save" type="submit" id="Save" value="Save This Report"></td>
    </tr>
  </form>
  <%
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->