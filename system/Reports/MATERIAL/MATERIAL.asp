<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")



fromdate=request("fromdate")

todate=request("todate")

pagepara="&fromdate="&fromdate&"&todate="&todate
where=""
if fromdate<>"" then

fromdate=dateadd("n",-60,fromdate)
end if
if todate<>"" then
todate=dateadd("n",-60,todate)
end if

  sqlA="select count(*) as  qty from pvs_tobps where ams_measuredatetime  between '"&fromdate&"' and '"&todate&"'"
  
  rs.open sqlA,conn,1,3
  
    if not rs.eof then
	qty=rs("qty")
	end if
	rs.close
	if clng(qty) > 50000 then
	
	response.Write("<script>window.alert('Query data is not greater than 50000!');location.href='material.asp';</script>")
	
	end if

pagename="/Reports/MATERIAL/MATERIAL.asp"


'sql="select * from  (select * from vw_pvstobps where ams_measuredatetime  between '"&fromdate&"' and '"&todate&"') k left join vw_material_lot m on (k.serialnumber=m.code)"

'sql="select * from (select * from (select * from vw_pvstobps where AMS_MEASUREDATETIME between '"&fromdate&"' and  '"&todate&"' )K left join job_2d_code m  on( k.serialnumber=m.code ))kk left join vw_MATERIAL_LOT on (kk.job_number=vw_MATERIAL_LOT.job_number and kk.sheet_number=vw_MATERIAL_LOT.subjobnumber)"

sql="select * from (select * from (select * from vw_pvstobps where AMS_MEASUREDATETIME between '"&fromdate&"' and '"&todate&"' )K left join job_2d_code m  on( k.serialnumber=m.code ))kk "

sql=sql + "left join (select job_number,subjobnumber,max(decode(PART_TYPE,'POT',MATERIAL_LOT_NUMBER)) POT,max(decode(PART_TYPE,'TOP',MATERIAL_LOT_NUMBER)) TOP,max(decode(PART_TYPE,'BOTTOM',MATERIAL_LOT_NUMBER)) BOTTOM,max(decode(PART_TYPE,'FRAME',MATERIAL_LOT_NUMBER)) FRAME from(select b.job_number,b.subjobnumber,b.labelid,c.MATERIAL_PART_NUMBER,c.material_lot_number, (CASE instr(upper(nvl(t2.description,0)), 'TOP') WHEN 0 THEN '' ELSE 'TOP' END ) || CASE instr(upper(nvl(t2.description,0)), 'BOTTOM') WHEN 0 THEN '' ELSE 'BOTTOM' END || CASE instr(upper(nvl(t2.description,0)), 'POT') WHEN 0 THEN '' ELSE 'POT' END || CASE instr(upper(nvl(t2.description,0)), 'FRAME') WHEN 0 THEN '' ELSE 'FRAME' END PART_TYPE from material_count_record b left join mr_dispatch c on(b.labelid=c.label_no) left join product_model t2 on (c.material_part_number = t2.item_name) ) group by job_number,subjobnumber  ) MATERIAL_LOT on (kk.job_number=MATERIAL_LOT.job_number and kk.sheet_number=MATERIAL_LOT.subjobnumber) "











'response.Write(sql)
'response.End()
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

<script language="javascript" type="text/javascript" src="../../Components/My97DatePicker/WdatePicker.js"></script>
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script language="javascript">
function retestcheck(ob,thisvalue)
{
	if (ob.checked)
	{
	location.href="Retest_Check.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
	else
	{
	location.href="Retest_Uncheck.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
}
</script>
<!--#include virtual="/Language/Reports/MATERIAL/Lan_MATERIAL.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/Reports/MATERIAL/MATERIAL.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td><span id="inner_SearchStoreTime"></span></td>
    <td><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value=""  size="30" readonly    >
    <img onclick="WdatePicker({el:'fromdate',dateFmt:'yyyy-MM-dd HH:mm:ss',onpicked:pickedFunc})" src="../../Images/dynCalendar.gif" align="absmiddle" style="cursor:pointer"/>
<script>
function pickedFunc(){
 $dp.$('d523_y').value=$dp.cal.getP('y');
 $dp.$('d523_M').value=$dp.cal.getP('M');
 $dp.$('d523_d').value=$dp.cal.getP('d');
 $dp.$('d523_HH').value=$dp.cal.getP('H');
 $dp.$('d523_mm').value=$dp.cal.getP('m');
 $dp.$('d523_ss').value=$dp.cal.getP('s');
 }
</script>
  &nbsp;<span id="inner_SearchTo"></span>
  <input name="todate" type="text" id="todate"  size="30" readonly value="" >
 <img onclick="WdatePicker({el:'todate',dateFmt:'yyyy-MM-dd HH:mm:ss',onpicked:pickedFunc})" src="../../Images/dynCalendar.gif" align="absmiddle" style="cursor:pointer"/>

  &nbsp; </td>
    
    <td height="20"><span id="inner_SearchPartNumber"></span></td>
<td height="20"></td>
    <!--
    <td><span id="inner_SearchCheckStatus"></span></td>
    <td>&nbsp;</td>
    <td>OBI Status</td>
    <td><select name="OBI_Status" id="OBI_Status">
        <option value="">-- Status --</option>
        <option value="NEW" <%IF OBI_Status="NEW" THEN%>SELECTED<%END IF%>>NEW</option>
        <option value="PENDING" <%IF OBI_Status="PENDING" THEN%>SELECTED<%END IF%>>PENDING</option>
        <option value="SUCCESS" <%IF OBI_Status="SUCCESS" THEN%>SELECTED<%END IF%>>SUCCESS</option>
        <option value="FAILURE" <%IF OBI_Status="FAILURE" THEN%>SELECTED<%END IF%>>FAILURE</option>
        <option value="VALIDATION_ERROR" <%IF OBI_Status="VALIDATION_ERROR" THEN%>SELECTED<%END IF%>>VALIDATION_ERROR</option>
      </select>
    </td>
	-->    
    <td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">
      </td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="225" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="27" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('MATERIAL_Export.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="27"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_ADID"></span></div></td>  
<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_ProductName"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_LineName"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_BoxName"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_MeasureCount"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_AMSMeasureDateTime"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_TestDay"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Serialnumber"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Year"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Week"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_WeekDay"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_ProdDate"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_WeekNum"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_DeltaTime"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_DeltaDay"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_ADFAI"></span></div></td>  
<td class="t-t-Borrow"><div align="center"><span id="inner_ERROR"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_CriterionError"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_SubJobs"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Pot"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Top"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Bottom"></span></div></td>
<td class="t-t-Borrow"><div align="center"><span id="inner_Frame"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 

%>
	<tr>
	    <td height="20"><div align="center"><%=i%></div></td>
		<td><div align="center"><%=rs("ad_id")%></div></td>
		<td height="20"><div align="center"><%=rs("productname")%></div></td>
		<td height="20"><div align="center"><%=rs("linename")%></div></td>
		<td><div align="center"><%=rs("boxname")%></div></td>
		<td><div align="center"><%=rs("measurecount")%></div></td>
		<td><div align="center"><%=rs("AMS_MEASUREDATETIME")%></div></td>
		<td><div align="center"><%=rs("TEST_DAY")%></div></td>
		<td><div align="center"><%=rs("serialnumber")%></div></td>
		<td><div align="center"><%=rs("year")%></div></td>
		<td><div align="center"><%=rs("week")%></div></td>
		<td><div align="center"><%=rs("weekday")%></div></td>
		<td><div align="center"><%=rs("MEASUREDATETIME")%></div></td>
		<td><div align="center"><%=rs("weeknum")%></div></td>
		<td><div align="center"><%=rs("delta_time")%></div></td>
        <td><div align="center"><%=rs("delta_day")%></div></td>
        <td><div align="center"><%=rs("line")%></div></td>
        <td><div align="center"><%=rs("adfail")%></div></td>
        <td><div align="center"><%=rs("error")%>&nbsp;</div></td>
        <td><div align="center"><%=rs("criterionerror")%>&nbsp</div></td>
        <td><div align="center"><%=rs("job_number")%></div></td>
        <td><div align="center"><%=rs("subjobnumber")%></div></td>
        <td><div align="center"><%=rs("POT")%></div></td>
        <td><div align="center"><%=rs("TOP")%></div></td>
         <td><div align="center"><%=rs("Bottom")%></div></td>
        <td><div align="center"><%=rs("Frame")%></div></td>
  </tr>
  
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="27"><div align="center"><span id="inner_NoRecords"></span></div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->