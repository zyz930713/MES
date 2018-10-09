<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
reportname=trim(request("reportname"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by FL.REPORT_TIME desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if reportname<>"" then
where=where&" and FL.FINAL_PARTYIELD_NAME='"&reportname&"'"
end if
if fromdate<>"" then
where=where&" and FL.REPORT_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and FL.REPORT_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&reportname="&reportname&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/FinalYield/FinalPartYield/FinalPartYieldList.asp"
SQL="select FL.*,F.FACTORY_NAME,U.USER_NAME from FINAL_PARTYIELD_LIST FL inner join FACTORY F on FL.FACTORY_ID=F.NID left join USERS U on FL.CREATOR_CODE=U.USER_CODE where FL.NID is not null "&where&order
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript">
function form2check()
{
var flag=0;
	with (document.form2)
	{
		if (factory.selectedIndex==0)
		{
		alert("You must select one factory");
		return false;
		}
		if (partselect.value=="m")
		{
			for (var i=0; i<part.length;i++)
			{
				if (part.options[i].selected==true)
				{
    			flag=flag+1;
				}
    		}
			if(flag==0||part.selectedIndex==0)
			{
			alert("Part selection can not be blank!");
			return false;
			}
		}
		else
		{
			if(parttext.value=="")
			{
			alert("Part input can not be blank!")
			return false;
			}
		}
	}
}
</script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<form name="form1" method="post" action="/Reports/FinalYield/FinalPartYield/FinalPartYieldList.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span>Search Final Part Yield Reports </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Report Name </span> </td>
    <td height="20"><input name="reportname" type="text" id="reportname" value="<%=reportname%>"></td>
    <td>Report Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate2" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate2" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#FFFFFF">
  <form action="/Reports/FinalYield/FinalPartYield/GenerateFinalPartYield.asp" method="post" name="form2" target="_self" onSubmit="return form2check()">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy">Generating conditions</td>
    </tr>
  <tr>
    <td>Factory: </td>
    <td><select name="factory">
	<option value="">Select</option>
	<%=getFactory("OPTION","",factorywhereinside,"","")%>
	</select></td>
    <td>Part:</td>
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
		  	<tbody id="tabSelect" style="display:">
              <tr>
                <td><select name="part" id="part">
                    <option value="">--Select--</option>
                    <%FactoryRight "M."%>
                    <%=getModel("OPTION",null,factorywhereoutsideand&" and rownum<=100"," order by M.ITEM_NAME","",idcount)%>
                  </select></td>
              </tr>
			  </tbody>
			  <tbody id="tabInput" style="display:none">
              <tr>
                <td><input name="parttext" type="text" id="parttext"></td>
              </tr>
			  </tbody>
            </table>
            </td>
          <td><img src="/Images/SpaceUp.gif" alt="Click to expend selection" width="14" height="8" onClick="javascript:if(document.all.part.multiple==true){document.all.part.size=document.all.part.size+3}"></td>
          <td rowspan="2"><input name="PartType" type="button" id="PartType" value="Multi" onClick="javascript:if(document.all.part.multiple==true){document.all.PartType.value='Multi';document.all.part.multiple=false}else{document.all.PartType.value='One';document.all.part.multiple=true}"></td>
          <td colspan="5" rowspan="2"><input name="partselect" type="hidden" id="partselect" value="m">
          <input name="PartType2" type="button" id="PartType2" value="Input" onClick="javascript:if(document.all['tabSelect'].style.display==''){document.all.PartType.disabled=true;document.all.PartType2.value='Select';document.all.partselect.value='i';document.all['tabSelect'].style.display='none';document.all['tabInput'].style.display=''}else{document.all.PartType.disabled=false;document.all.PartType2.value='Input';document.all.partselect.value='m';document.all['tabSelect'].style.display='';document.all['tabInput'].style.display='none'}"></td>
          </tr>
        <tr>
          <td><img src="/Images/SpaceDown.gif" alt="Click to decrease selection" width="14" height="8" onClick="javascript:if(document.all.part.multiple==true&&document.all.part.size>3){document.all.part.size=document.all.part.size-3}"></td>
          </tr>
      </table>
      </td>
    <td>Job Created Time</td>
    <td>From
      <input name="jfromdate" type="text" id="jfromdate" value="<%=dateadd("d",-7,date())%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.jfromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
      <input name="jfromhour" type="text" id="jfromhour" value="00:00" size="5">
      &nbsp;to
<input name="jtodate" type="text" id="jtodate" value="<%=date()%>" size="10">
<script language=JavaScript type=text/javascript>
	</script>
<script language=JavaScript type=text/javascript>function calendar4Callback(date, month, year)
	{
	document.all.jtodate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
                        </script>
&nbsp; <input name="jtohour" type="text" id="jtohour" value="00:00" size="5"></td>
  </tr>
  <tr>
    <td height="20" colspan="6"><input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Generate Final Part Yield Report at once"></td>
    </tr>
  </form>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="48%">Browse Final Part Yield Report </td>
        <td width="52%"><div align="right">User:
          <% =session("User") %>
        </div></td>
      </tr>
    </table></td>
  </tr>
  
    
  <tr>
    <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
	<td class="t-t-Borrow"><div align="center">Delete</div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WIP_NAME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WIP_NAME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center">Period</div></td>
    <td class="t-t-Borrow"><div align="center">Factory</div></td>
    <td class="t-t-Borrow"><div align="center">Creator</div></td>
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
      </div></td>
    <td><div align="center"><%if session("code")=rs("CREATOR_CODE") then%><span class="red" style="cursor:hand" onClick="javascript:location.href='DeleteFinalPartYieldReport.asp?id=<%=rs("NID")%>&finalPart_name=<%=rs("FINAL_PartYIELD_NAME")%>&path=<%=path%>&query=<%=query%>'"><img src="/Images/IconDelete.gif"></span><%else%><img src="/Images/IconDelete_No.gif"><%end if%></div></td>
    <td height="20">
      <div align="center">
        <% =formatdate(rs("REPORT_TIME"),application("longdateformat"))%>
      </div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('FinalPartYieldReport.asp?finalPart_id=<%=rs("NID")%>&finalPart_name=<%=rs("FINAL_PartYIELD_NAME")%>&finalPart_report_time=<%=rs("REPORT_TIME")%>&from_time=<%=rs("FROM_TIME")%>&to_time=<%=rs("TO_TIME")%>&factory_id=<%=rs("FACTORY_ID")%>')"><%= rs("FINAL_PartYIELD_NAME") %></span>&nbsp;</div></td>
    <td><div align="center"><% =formatdate(rs("FROM_TIME"),application("longdateformat"))%> to <% =formatdate(rs("TO_TIME"),application("longdateformat")) %> </div></td>
    <td><div align="center"><% =rs("FACTORY_NAME") %></div></td>
    <td><div align="center"><% =rs("USER_NAME") %></div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
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