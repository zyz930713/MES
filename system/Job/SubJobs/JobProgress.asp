<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
jobstatus=trim(request("jobstatus"))
currentstation=trim(request("currentstation"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if jobstatus<>"" then
where=where&" and J.STATUS="&jobstatus
else
where=where&" and J.STATUS=0"
end if
if currentstation<>"" then
where=where&" and J.CURRENT_STATION_ID='"&currentstation&"'"
end if

if request("thisdate")<>"" then
thisdate=cdate(request("thisdate"))
else
thisdate=date()
end if
if request("thistime")<>"" then
thistime=cdate(request("thistime"))
else
thistime=#8:00#
end if
if request("hour_select")<>"" then
hour_select=cint(request("hour_select"))
else
hour_select=18
end if
current=hour_select

set rs1=server.CreateObject("ADODB.RECORDSET")
pagename="/Job/JobProgress.asp"
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&thisdate="&thisdate&"&thistime="&thistime&"&hour_select="&hour_select
SQL="SELECT 1,J.JOB_NUMBER,J.SHEET_NUMBER,J.STATUS,J.CURRENT_STATION_ID,P.STATIONS_INDEX FROM JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where ((J.START_TIME>=to_date('"&thisdate&" 00:00:00','yyyy-mm-dd  hh24:mi:ss') and J.START_TIME<=to_date('"&thisdate&" 23:59:59','yyyy-mm-dd  hh24:mi:ss')) or (J.CLOSE_TIME>=to_date('"&thisdate&" 00:00:00','yyyy-mm-dd  hh24:mi:ss') and J.CLOSE_TIME<=to_date('"&thisdate&" 23:59:59','yyyy-mm-dd  hh24:mi:ss')))"&where&order
session("SQL")=SQL
rs.open SQL,conn,1,3
%>				
<html>
<head>
<title>Barcode System</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language=JavaScript src="/Components/sniffer.js" type=text/javascript></script>
<script language=JavaScript src="/Components/dynCalendar.js" type=text/javascript></script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="javascript">
function time_select()
{
location.href="<%=pagename%>?jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&thisdate=<%=thisdate%>&thistime=<%=thistime%>&hour_select="+document.all.selectforhour.options[document.all.selectforhour.selectedIndex].value
}
</script>
</head>

<body onLoad="language_jobnote()">
<form action="/Job/SubJobs/JobProgress.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span>Search Job </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Job Number</span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td>Sheet Number </td>
    <td><input name="sheetnumber" type="text" id="sheetnumber" value="<%=sheetnumber%>"></td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    </tr>
  <tr>
    <td height="20">Status</td>
    <td height="20"><select name="jobstatus" id="jobstatus">
      <option value="">All</option>
      <option value="0" <%if jobstatus="0" then%>selected<%end if%>>Opened</option>
      <option value="2" <%if jobstatus="2" then%>selected<%end if%>>Paused</option>
      <option value="1" <%if jobstatus="1" then%>selected<%end if%>>Closed</option>
      <option value="3" <%if jobstatus="3" then%>selected<%end if%>>Locked</option>
    </select></td>
    <td>Current Station </td>
    <td><select name="currentstation" id="currentstation">
      <option value="">All</option>
	  <%FactoryRight "S."%>
      <%=getStation(true,"OPTION",currentstation,factorywhereoutsideand,"","","")%>
    </select></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    <td>&nbsp;</td>
    </tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#3399FF" bordercolordark="#FFFFFF">
  
  <tr valign="middle" class="t-c-greenCopy">
    <td height="20" colspan="<%=current+4%>" align="center"><div align="left">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="t-c-greenCopy">Progress of Job </td>
          <td><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobProgress_Export.asp?thisdate=<%=thisdate%>&thistime=<%=thistime%>&hour_select=<%=hour_select%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
        </tr>
      </table> 
      </div>
    <div align="right"></div></td>
  </tr>
  <tr valign="middle" class="t-t-FIN" >
    <td rowspan="2" align="center" nowrap><div align="center">No.</div></td>
    <td rowspan="2" align="center" nowrap><div align="center">Status</div></td>
    <td rowspan="2" align="center" nowrap><div align="center">Finished</div></td>
    <td rowspan="2" align="center" nowrap><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc<%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc<%=pagepara%>'"></div></td>
    <td colspan="<%=current%>" align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><%if thistime>#0:00# then%><a href="<%=pagename%>?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&thisdate=<%=thisdate%>&thistime=<%=dateadd("n",-30,thistime)%>&hour_select=<%=hour_select%>"><img src="/Images/prev.gif" width="10" height="10" border="0"></a><%end if%>Before 
          <select name="o1" onChange="javascript:location.href='<%=pagename%>?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&thisdate=<%=thisdate%>&thistime='+this.options[this.selectedIndex].value+'&hour_select=<%=hour_select%>'">
		  <%o1=#0:0#
		  for h1=0 to 47%>
		  <option value="<%=o1%>" <%if o1=thistime then%>selected<%end if%>><%=o1%></option>
		  <%
		  if o1=thistime then exit for
		  o1=dateadd("n",30,o1)
		  next%>
		  </select>
          </td>
        <td height="20"><div align="center">
            On 
            <input name="thisdate" type="text" value="<%=thisdate%>" size="10" readonly="true">
            <script language=JavaScript type=text/javascript>
	function calendar11Callback(date, month, year)
	{
	var strdate=year + '-' + month + '-' + date
	document.all.thisdate.value=strdate
	location.href="<%=pagename%>?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&thisdate="+strdate+"&thistime=<%=thistime%>&hour_select=<%=hour_select%>"
	}
    calendar11 = new dynCalendar('calendar11', 'calendar11Callback');
	</script>
&nbsp;
<select name="selectforhour" onChange="time_select()">
  <option>Select Hours</option>
  <%
			  for i=2 to 48 step 2
			  %>
  <option value="<%=i%>" <%if i=hour_select then%>selected<%end if%>><%=i/2%> Hour</option>
  <%
			  next
			  %>
</select>
        </div></td>
        <td><div align="right">
		<%
		ntime=thistime
		for i=1 to current-1
		ntime=dateadd("n",30,ntime)
		if ntime>#23:59# then exit for  
		next
		%>
		After 
		<%= ntime %><%if ntime<#23:30# then%><a href="<%=pagename%>?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&thisdate=<%=thisdate%>&thistime=<%=dateadd("n",30,ntime)%>&hour_select=<%=hour_select%>"><img src="/Images/next.gif" width="10" height="10" border="0"></a><%end if%></div></td>
      </tr>
    </table></td>
  </tr>
  <tr valign="middle" class="t-t-FIN">
    <%
					  ktime=thistime
					  for i=1 to current %>
    <td align="center"><div align="left"><img src="/Images/Split1_blue.gif" width="3" height="20" align="absmiddle">
      <%
					  sktime=cstr(ktime)
					  if len(sktime)>7 then
					  strktime=left(sktime,5)
					  else
					  strktime=left(sktime,4)
					  end if
					  %>
          <%=strktime%></div></td>
    <%
					  ktime=dateadd("n",30,ktime)
					  if ktime>#23:59# then exit for
					  next%>
  </tr>
<% 
if not rs.eof then
WW=1
while not rs.eof 
%>
  <tr valign="middle">
    <td height="20" align="center"><div align="center">
      <% =WW%>    
    </div></td>
    <td align="center" nowrap><div align="center">
      <%
	  if rs("STATUS")="0" then
	  simg="Opened"
	  aimg="Pause,Abort"
	  alt="pause,abort"
	  apage="PauseJob,AbortJob"
	  elseif rs("STATUS")="1" then
	  simg="Closed"
	  aimg=""
	  alt=""
	  apage=""
	  elseif rs("STATUS")="2" then
	  simg="Paused"
	  aimg="Start"
	  alt="Start"
	  apage="StartJob"
	  elseif rs("STATUS")="3" then
	  simg="Locked"
	  aimg="Start"
	  alt="Unlock"
	  apage="UnlockJob"
	  elseif rs("STATUS")="4" then
	  simg="Aborted"
	  aimg=""
	  alt=""
	  apage=""
	  end if%>
      <img src="/Images/<%=simg%>.gif"></div></td>
    <td align="center" nowrap>
	  
      <div align="center">
	    <%
	astations=split(rs("STATIONS_INDEX"),",")
	totalstations=ubound(astations)
	for s=0 to ubound(astations)
		if astations(s)=rs("CURRENT_STATION_ID") then
		cstation=s
		end if
	next
	if totalstations<>0 then
	progress=cstation/totalstations
	else
	progress=0
	end if%>
	    <%=formatpercent(progress,2,-1)%>&nbsp;</div></td>
    <td align="center" nowrap><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>" target="_blank"><%=rs("JOB_NUMBER")%>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></a></div></td>
    <%
	j=0
	id=""
	beforeid=""
	cid=""
	title=""
	beforetitle=""
	ctitle=""
	color=""
	beforecolor=""
	ccol=""
	code=""
	beforecode=""
	ccode=""
	start_time=""
	before_start_time=""
	c_start_time=""
	close_time=""
	before_close_time=""
	c_close_time=""
	cspan=""
	tdout=""
	ktime=thistime
						  
		for i=1 to current
		if ktime>=#23:30# then
		nexttime=#0:0#
		kdate=dateadd("d",1,thisdate)
		else
		nexttime=dateadd("n",30,ktime)
		kdate=thisdate
		end if
			SQL1="select 1,JS.STATION_ID,JS.START_TIME,JS.CLOSE_TIME,JS.OPERATOR_CODE,S.STATION_NAME from JOB_STATIONS JS inner join STATION S on JS.STATION_ID=S.NID where JS.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and JS.SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and ((JS.START_TIME>=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.START_TIME<=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')) or (JS.CLOSE_TIME>=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME<=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')) or (JS.START_TIME<=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME>=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')))"
			rs1.open SQL1,conn,1,3
			
			if not rs1.eof then
			id=rs1("STATION_ID")
			title=rs1("STATION_NAME")
			code=rs1("OPERATOR_CODE")
			start_time=rs1("START_TIME")
			close_time=rs1("CLOSE_TIME")
				if rs1("STATION_ID")=rs("CURRENT_STATION_ID") then
				color="#FF9B9B"
				else	
				color="#B6EAB0"
				end if
			else
			id=""
			title="&nbsp;"
			code="&nbsp;"
			start_time="&nbsp;"
			close_time="&nbsp;"
			color="#FFFFFF"
			end if
			rs1.close
			
			if id=beforeid and beforeid<>"" then
			j=j+1
			else
				if i<>1 then
				cid=cid&beforeid&","
				ccode=ccode&beforecode&","
				ctitle=ctitle&beforetitle&","
				ccol=ccol&beforecolor&","
				c_start_time=c_start_time&before_start_time&","
				c_close_time=c_close_time&before_close_time&","
				cspan=cspan&j&","
				end if
			j=1
			end if

		beforeid=id
		beforecode=code
		before_start_time=start_time
		before_close_time=close_time
		beforetitle=title
		beforecolor=color
		ktime=dateadd("n",30,ktime)
		lastcurrent=i
		if ktime>#23:59# then exit for  
		next
		
		cid=cid&beforeid&","
		ccode=ccode&beforecode&","
		c_start_time=c_start_time&before_start_time&","
		c_close_time=c_close_time&before_close_time&","
		ctitle=ctitle&beforetitle&","
		ccol=ccol&color&","
		cspan=cspan&j&","
		
		cid=left(cid,len(cid)-1)
		ccode=left(ccode,len(ccode)-1)
		c_start_time=left(c_start_time,len(c_start_time)-1)
		c_close_time=left(c_close_time,len(c_close_time)-1)
		ctitle=left(ctitle,len(ctitle)-1)
		ccol=left(ccol,len(ccol)-1)
		cspan=left(cspan,len(cspan)-1)
		cida=split(cid,",")
		ccodea=split(ccode,",")
		c_start_time_a=split(c_start_time,",")
		c_close_time_a=split(c_close_time,",")
		ctitlea=split(ctitle,",")
		ccola=split(ccol,",")
		cspana=split(cspan,",")
		
		p=0
		q=1
		while q<=lastcurrent
		tdout=tdout&"<td colspan='"&cspana(p)&"' align=center bgcolor='"&ccola(p)&"' title='"&ctitlea(p)&" from "&c_start_time_a(p)&" to "&c_close_time_a(p)&"("&ccodea(p)&")'>"
			if ctitlea(p)<>"&nbsp;" and len(ctitlea(p))>cspana(p)*4 then
			tdout=tdout&left(ctitlea(p),cspana(p)*4)
			else
			tdout=tdout&ctitlea(p)
			end if
		
		tdout=tdout&"</td>"
		q=q+cspana(p)
		p=p+1
		wend%><%=tdout%>  </tr>
<% 
rs.MoveNext
WW=WW+1
Wend
else%>
  <tr valign="middle">
    <td colspan="<%=current%>" align="center">No Records </td>
  </tr>
<%
end if
%>
</table>
<!--#include virtual="/Components/JobNote.asp" -->
</body>
</html>