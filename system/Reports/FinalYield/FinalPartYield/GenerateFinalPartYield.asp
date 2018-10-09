<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
create_time=now()
rnd_key=gen_key(10)
thiserror=""

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=trim(request.Form("factory"))
if request.Form("partselect")="m" then
part_number=replace(trim(request.Form("part"))," ","")
else
part_number=ucase(trim(request.Form("parttext")))
end if
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
pagename="/Reports/FinalYield/FinalPartYiled/GenerateFinalPartYield.asp"

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="FINAL_PART_YIELD"
cmd.CommandType=4 
'response.Write(factory&";"&part_number&";"&part_number&";"&fromtime&";"&totime&";"&session("code")&";"&rnd_key)
'response.End()
cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 10, factory)
cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 4000, part_number)
cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, fromtime)
cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, totime)
cmd.Parameters.Append cmd.CreateParameter("creator_code", adVarChar, adParamInput, 4, session("code"))
cmd.Parameters.Append cmd.CreateParameter("rnd_key", adVarChar, adParamInput, 10, rnd_key)
cmd.execute
set cmd=nothing
if err.number=0 then
word="Fail to create schedule.\n创建计划任务失败！"
end if
FactoryRight "S."
SQL="select 1,Y.* from FINAL_PartYIELD_DETAIL_TEMP Y where Y.CREATOR_CODE='"&session("code")&"' and Y.RND_KEY='"&rnd_key&"' order by Y.Part_NAME"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="javascript">
function formcheck()
{
	with(document.form1)
	{
		if(finalpart_name.value=="")
		{
		alert("Name of report cannot be blank");
		return false;
		}
	}
}
</script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">Browse Final Yield (from <%=fromtime%> to <%=totime%>) </td>
  </tr>
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">User:
    <% =session("User") %></td>
  </tr>
  <tr>
    <td height="20" colspan="7">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Part Name </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Assembly Input </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Assembly Output </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">First Passed  Yield </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Stocked Output </div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
  </tr>
  <%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalPartYield/FinalPartDetail.asp" method="post" name="formDetail" target="_blank">
<%
while not rs.eof
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.store_nids.value='<%=rs("STORE_NIDS")%>';document.all.part_name.value='<%=rs("PART_NAME")%>';document.formDetail.submit()"><%=rs("PART_NAME")%></span></td>
    <td height="20"><div align="center"><%=rs("ASSEMBLY_INPUT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("ASSEMBLY_OUTPUT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=rs("OUTPUT_QUANTITY")%></div></td>
    <td height="20" class="<%=yieldcolor%>"><div align="center"><%=formatpercent(csng(rs("FINAL_YIELD")),2,-1)%></div></td>
    </tr>
  <%
rs.movenext
i=i+1
wend
%>
<input name="store_nids" type="hidden" value="">
<input name="part_name" type="hidden" value="">
</form>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="7">&nbsp;</td>
  </tr>
  <form name="form1" method="post" action="/Reports/FinalYield/FinalPartYield/SaveFinalPartYield.asp" onSubmit="return formcheck()">
    <tr>
      <td height="20" colspan="7">Generating time:
        <% =formatdate(create_time,application("longdateformat"))%>
      &nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="7"> Report name:
        <input type="checkbox" name="checkbox" value="checkbox">
        save 
        <input name="finalpart_name" type="text" id="finalpart_name">
          <input name="from_time" type="hidden" id="from_time" value="<%=fromtime%>">
          <input name="to_time" type="hidden" id="to_time" value="<%=totime%>">
<input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
          <input name="factory" type="hidden" id="factory" value="<%=factory%>">
      <input name="Save" type="submit" id="Save" value="Save This Report">      </td>
    </tr>
  </form>
  <%
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->