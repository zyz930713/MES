<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
server.ScriptTimeout=500%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
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
factory=trim(request.Form("factory"))
fromdate=request.Form("jfromdate")
todate=request.Form("jtodate")
fromhour=request.Form("jfromhour")
tohour=request.Form("jtohour")
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
pagename="/Reports/FinalYield/FinalSeriesYiled/GenerateFinalSeriesYield.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="FINAL_SERIES_YIELD"
cmd.CommandType=4 
'response.Write(factory&":"&fromtime&":"&totime&":"&session("code")&":"&rnd_key)
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
SQL="select TARGET_YIELD,TARGET_FIRSTYIELD from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_yield=rs("TARGET_YIELD")
factory_target_firstyield=rs("TARGET_FIRSTYIELD")
end if
rs.close
FactoryRight "S."
SQL="select * from FINAL_SERIESYIELD_DETAIL_TEMP where CREATOR_CODE='"&session("code")&"' and RND_KEY='"&rnd_key&"' order by SERIES_NAME"
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
    <td height="20" colspan="11" class="t-c-greenCopy">Browse Final Series Yield (from <%=fromtime%> to <%=totime%>) </td>
  </tr>
  <tr>
    <td height="20" colspan="11" class="t-c-greenCopy">User:
    <% =session("User") %></td>
  </tr>
  <tr>
    <td height="20" colspan="11">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Series Name </div></td>
    <td class="t-t-Borrow"><div align="center">Assembly Input </div></td>
    <td class="t-t-Borrow"><div align="center">Assembly Output </div></td>
    <td class="t-t-Borrow"><div align="center">Stocked Output </div></td>
    <td class="t-t-Borrow"><div align="center">First Passed  Yield </div></td>
    <td class="t-b-newyear"><div align="center">First Target Yield </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
    <td class="t-b-newyear"><div align="center">Inernal Yield </div></td>
    <td height="20" class="t-b-newyear"><div align="center">Target Yield</div></td>
    <td class="t-t-Borrow"><div align="center">Delta</div></td>
  </tr>
  <%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalSeriesYield/FinalSeriesDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_assembly_input=0
total_assembly_output=0
total_final_output=0
while not rs.eof
%>
  <tr <%if rs("OVERALL_EXCEPTION")="1" then%>class="today"<%end if%>>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.store_nids.value='<%=rs("STORE_NIDS")%>';document.all.series_name.value='<%=rs("SERIES_NAME")%>';document.formDetail.submit()"><%=rs("SERIES_NAME")%></span></td>
    <td><div align="center"><%=rs("ASSEMBLY_INPUT_QUANTITY")%></div></td>
    <td><div align="center"><%=rs("ASSEMBLY_OUTPUT_QUANTITY")%></div></td>
    <td><div align="center"><%=rs("OUTPUT_QUANTITY")%></div></td>
    <td class="<%'=getYieldColor(rs("ASSEMBLY_YIELD"),rs("SERIES_TARGET_FIRSTYIELD"))%>"><div align="center"><%=formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_FIRSTYIELD"))/100,2,-1)%></div></td>
    <td height="20" class="<%=getYieldColor(rs("FINAL_YIELD"),rs("SERIES_TARGET_YIELD"))%>"><div align="center"><%=formatpercent(csng(rs("FINAL_YIELD")),2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_INTERNALYIELD")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_YIELD"))/100,2,-1)%></div></td>
    <td><div align="center">
	<%if csng(rs("FINAL_YIELD"))<>0 then
	delta=csng(rs("FINAL_YIELD"))-csng(rs("SERIES_TARGET_YIELD"))/100
	else
	delta=0
	end if%><%=formatpercent(delta,2,-1)%></div></td>
  </tr>
  <%
	if rs("OVERALL_EXCEPTION")="0" then
	total_assembly_input=total_assembly_input+csng(rs("ASSEMBLY_INPUT_QUANTITY"))
	total_assembly_output=total_assembly_output+csng(rs("ASSEMBLY_OUTPUT_QUANTITY"))
	total_final_output=total_final_output+csng(rs("OUTPUT_QUANTITY"))
	end if
rs.movenext
i=i+1
wend
if total_assembly_input<>0 then
total_assembly_yield=total_assembly_output/total_assembly_input
total_final_yield=total_final_output/total_assembly_input
else
total_assembly_yield=0
total_final_yield=0
end if
%> 
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td height="20"><div align="center"><%=total_assembly_input%></div></td>
    <td height="20"><div align="center"><%=total_assembly_output%></div></td>
    <td height="20"><div align="center"><%=total_final_output%></div></td>
    <td height="20" class="<%=getYieldColor(total_assembly_yield,factory_target_firstyield)%>"><div align="center">&nbsp;<%=formatpercent(total_assembly_yield,2,-1)%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(factory_target_yield)/100,2,-1)%></div></td>
    <td height="20" class="<%=getYieldColor(total_final_yield,factory_target_yield)%>"><div align="center"><%=formatpercent(total_final_yield,2,-1)%></div></td>
    <td>&nbsp;</td>
    <td height="20"><div align="center"><%=formatpercent(csng(factory_target_yield)/100,2,-1)%></div></td>
    <td><div align="center">
      <%if csng(total_final_yield)<>0 then
	delta=csng(total_final_yield)-csng(factory_target_yield)/100
	else
	delta=0
	end if%>
      <%=formatpercent(delta,2,-1)%></div></td>
  </tr>
<input name="store_nids" type="hidden" value="">
<input name="series_name" type="hidden" value="">
</form>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="11"><table cellspacing="0" cellpadding="0">
      <tr height="18">
        <td width="78" height="18" class="t-b-Green">&nbsp;</td>
        <td colspan="7" width="350">means cum yield meet target</td>
      </tr>
      <tr height="18">
        <td height="18" class="t-b-Yellow"></td>
        <td colspan="7">means delta is between 0% to 0.5%</td>
      </tr>
      <tr height="21">
        <td height="21" class="t-b-Red"></td>
        <td colspan="7">means delta is bigger than 0.5%</td>
      </tr>
    </table></td>
  </tr>
  <form name="form1" method="post" action="/Reports/FinalYield/FinalSeriesYield/SaveFinalSeriesYield.asp" onSubmit="return formcheck()">
    <tr>
      <td height="20" colspan="11">Generating time:
        <% =formatdate(create_time,application("longdateformat"))%>
      &nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="11">Chart Report:
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
      <td height="20" colspan="11"> Report name:
        <input name="finalseries_name" type="text" id="finalseries_name">
          <input name="factory_target_yield" type="hidden" id="factory_target_yield" value="<%=factory_target_yield%>">
          <input name="factory_target_firstyield" type="hidden" id="factory_target_firstyield" value="<%=factory_target_firstyield%>">
          <input name="from_time" type="hidden" id="from_time" value="<%=fromtime%>">
          <input name="to_time" type="hidden" id="to_time" value="<%=totime%>">
<input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
          <input name="factory" type="hidden" id="factory" value="<%=factory%>">
          <input name="total_assembly_input" type="hidden" id="total_assembly_input" value="<%=total_assembly_input%>">
          <input name="total_assembly_output" type="hidden" id="total_assembly_output" value="<%=total_assembly_output%>">
          <input name="total_assembly_yield" type="hidden" id="total_assembly_yield" value="<%=total_assembly_yield%>">
          <input name="total_final_output" type="hidden" id="total_final_output" value="<%=total_final_output%>">
          <input name="total_final_yield" type="hidden" id="total_final_yield" value="<%=total_final_yield%>">
      <input name="Save" type="submit" id="Save" value="Save This Report">      </td>
    </tr>
  </form>
  <%
else
%>
  <tr>
    <td height="20" colspan="11"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->