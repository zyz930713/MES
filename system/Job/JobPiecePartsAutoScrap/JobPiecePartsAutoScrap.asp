<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
	MaterialPartNumber=request("txtMaterialPartNumber")
	LineName=request("txtLineName")
 
	fromdate=request("fromdate")
	fromtime=request("fromtime")
	todate=request("todate")
	totime=request("totime")
	time0=now   
	if isnull(fromdate) or fromdate=""  then
		fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
	end if
	
	if isnull(fromtime) or fromtime=""  then
		fromtime="14:30:00"
	end if
	todate=request("todate")
	if isnull(todate) or todate=""  then
		todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
	end if
	totime=request("totime")
	if isnull(totime) or totime="" then
		totime="14:30:00"
	end if
	isSave=false
	if(request.QueryString("Action")="2") then
		SQLStr="SELECT  a.job_number, a.subjobnumber,b.material_part_number, a.AMOUNT as comsume_material_qty, "
		SQLStr=SQLStr+" (select js.station_start_quantity*GET_BOM_RATIO(b.job_number, b.material_part_number)"
		SQLStr=SQLStr+" from job_stations js, station s"
		SQLStr=SQLStr+" where js.station_id=s.nid and js.job_number=a.job_number"
		SQLStr=SQLStr+" and js.sheet_number=a.subjobnumber and s.mother_station_id=a.station_id) as material_start_qty,a.SCAN_DATETIME,j.line_name"
		SQLStr=SQLStr+" FROM MATERIAL_COUNT_RECORD a, mr_dispatch b,job j"
		SQLStr=SQLStr+" where a.labelid= b.label_no"
		SQLStr=SQLStr+" and j.job_number=a.job_number and j.sheet_number=a.subjobnumber"
		
		if(MaterialPartNumber<>"")then
			SQLStr=SQLStr+" AND b.material_part_number='"+MaterialPartNumber+"'"
		end if 
		if(LineName<>"")then
			SQLStr=SQLStr+" AND j.line_name='"+LineName+"'"
		end if
		if(fromdate<>"")then
			SQLStr=SQLStr+" AND scan_datetime>=to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		end if
		
		if(todate<>"")then
			SQLStr=SQLStr+" AND scan_datetime<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		end if
		SQLStr=SQLStr+" order by j.line_name,a.job_number ,a.subjobnumber,b.material_part_number,a.SCAN_DATETIME "
		rs.open SQLStr,conn,1,3
		isSave=false
	end if
	
  	if(request.QueryString("Action")="1") then
		rowno=request("txtCount")
		for i=1 to rowno
			Set TypeLib = CreateObject("Scriptlet.TypeLib")
 			Guid = TypeLib.Guid
			set rsSave=server.createobject("adodb.recordset")
			MATERIAL_PART_NUMBER=request("txtMaterialPartNumber"&i)	
			MaterialStartQty=request("txtMaterialStartQty"&i)
			ConsumeQty=request("txtComsumeMaterialQty"&i)
			ActualConsumeQty=request("txtActualComsumeMaterialQty"&i)
			LineName=request("txtLineName"&i)
			rsSave.open "select * from PIECE_PARTS_SCRAP where 1=2",conn,1,3
			rsSave.addnew
			rsSave("MATERIAL_PART_NUMBER")=MATERIAL_PART_NUMBER
			rsSave("SCRAP_OP_CODE")=session("code")
			rsSave("QTY")=ConsumeQty-MaterialStartQty
			rsSave("ACTUAL_QTY")=ActualConsumeQty-MaterialStartQty
			rsSave("LINE_NAME")=LineName
			
			rsSave("SCRAP_DATETIME")=now()
			rsSave("SCRAP_TYPE")="1"
			rsSave("IS_SENT_APPROVE")="0"
			rsSave("GUID")=Guid
			rsSave.update
			
			isSave=true
		next 
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript">
	function SearchData()
	{
		if(document.getElementById("txtLineName").value=="")
		{
			window.alert("请输入线别！");
			return;
		}
		
		form1.action="JobPiecePartsAutoScrap.asp?Action=2";
		form1.submit();
	}
	
	function SaveData()
	{
		form1.action="JobPiecePartsAutoScrap.asp?Action=1";
		form1.submit();
	}
	
</script>
</head>

<body>
<form method="post" name="form1" target="_self" >
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
   <td height="20" colspan="8" class="t-b-midautumn"> 
		报废查询
	</td>
  </tr>
 <tr>
   <tr>
    <td height="20">线别</td>
    <td height="20">
	<input type="text" name="txtLineName" id="txtLineName" value="<%=LineName%>">
	</td>
	 <td height="20">原材料型号</td>
    <td height="20" colspan="4">
	<input type="text" name="txtMaterialPartNumber" id="txtMaterialPartNumber" value="<%=MaterialPartNumber%>">
	</td>
  </tr>
  <Tr>
  	 <td>报废时间 从:</td> 
	 <td align="left"><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	<select name="fromtime" id="fromtime">
   			 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
</td>
<td align="left">到:</td>
<td colspan="4" align="left">
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
&nbsp;

<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="SearchData()">
&nbsp;
<%if isSave<>true then%>
<input type="button" name="btnSave" id="btnSave" value="保存" onclick="SaveData()"  class="t-b-Yellow">
<%end if %>


<input type="button" name="btnScrapSearch" id="btnScrapSearch" value="报废查询" onclick="window.open('JobPiecePartsAutoScrapSearch.asp')"  class="t-b-Yellow">


<td>
  </tr>
</table>

<table width="98%"  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
		<td height="20" class="t-t-Borrow">序号</td>
		<td height="20" class="t-t-Borrow">工单号</td>
		<td height="20" class="t-t-Borrow">子工单号</td>
		<td height="20" class="t-t-Borrow">原材料料号</td>
		<td height="20" class="t-t-Borrow">开始数量</td>
		<td height="20" class="t-t-Borrow">消耗数量</td>
		<td height="20" class="t-t-Borrow">实际消耗数量</td>
		<td height="20" class="t-t-Borrow">线别</td>
	</tr>
	<%if(request.QueryString("Action")="2") then 
		i=1
		while not rs.eof
		
	%>
		<tr>
			<td><%=i%></td>
			<td> <input type="hidden" name="txtJobNumber<%=i%>" id="txtJobNumber<%=i%>" value="<%=rs("job_number")%>"> <%=rs("job_number")%></td>
			<td> <input type="hidden" name="txtSheetNumber<%=i%>" id="txtSheetNumber<%=i%>" value="<%=rs("subjobnumber")%>"> <%=rs("subjobnumber")%></td>
			<td> <input type="hidden" name="txtMaterialPartNumber<%=i%>" id="txtMaterialPartNumber<%=i%>" value="<%=rs("MATERIAL_PART_NUMBER")%>"><%=rs("MATERIAL_PART_NUMBER")%> </td>
			<td> <input type="hidden" name="txtMaterialStartQty<%=i%>" id="txtMaterialStartQty<%=i%>" value="<%=rs("material_start_qty")%>"> <%=rs("material_start_qty")%></td>
			<td> <input type="hidden" name="txtComsumeMaterialQty<%=i%>" id="txtComsumeMaterialQty<%=i%>" value="<%=rs("comsume_material_qty")%>"> <%=rs("comsume_material_qty")%></td>
			<td> <input type="text" name="txtActualComsumeMaterialQty<%=i%>" id="txtActualComsumeMaterialQty<%=i%>" value="<%=rs("comsume_material_qty")%>"> </td>
			<td> <input type="hidden" name="txtLineName<%=i%>" id="txtLineName<%=i%>" value="<%=rs("line_name")%>"><%=rs("line_name")%> </td>
		 
		</tr>
	<%
			i=i+1
			rs.movenext
		wend
		end if %>
</table>
<input type="hidden" name="txtCount" id="txtCount" value="<%=i-1%>">
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->