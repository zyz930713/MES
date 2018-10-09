<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->


<%
	Family= request.querystring("Family") 
	Series= request.querystring("Series") 
	SubSeries= request.querystring("SubSeries") 
	Model= request.querystring("Model") 
	
	SQL= "select FA.FAMILY_NAME" 
	if Series<>"" then
		SQL=SQL+" ,SN.SERIES_NAME"
	end if
	
	if SubSeries<>"" then
		SQL=SQL+" ,SS.SUBSERIES_NAME"
	end if
	
	if Model<>"" then
		SQL=SQL+" ,TEMP.part_number_tag"
	end if
	
	SQL=SQL+"  ,temp.defect_name,sum(temp.defect_quantity) as Quantity"
	
	SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM,( "
	SQL=SQL+ "SELECT j.part_number_tag,df.defect_name,jd.defect_quantity "
	SQL=SQL+" FROM job_defectcodes jd,job_master_store j,defectcode df "
	SQL=SQL+"  WHERE jd.job_number = j.job_number  AND jd.defect_code_id = df.nid(+) "
	SQL=SQL+" AND j.store_time > to_date('"& session("FromDateTime") &"','yyyy-mm-dd hh24:mi:ss') "
	SQL=SQL+" AND j.store_time <= to_date('"& session("ToDateTime") &"','yyyy-mm-dd hh24:mi:ss') "
	SQL=SQL+" ) temp "
	SQL=SQL+" where temp.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID "
	
	if Family<>"" then
		SQL=SQL+" AND FA.FAMILY_NAME='"&Family&"'"
	END IF
	if Series<>"" then
		SQL=SQL+" and SN.SERIES_NAME='"&Series&"'"
	end if
	
	if SubSeries<>"" then
		SQL=SQL+" and SS.subseries_name='"&SubSeries&"'"
	end if
	
	if Model<>"" then
		SQL=SQL+" and TEMP.part_number_tag='"&Model&"'"
	end if
	
	SQL=SQL+" GROUP BY FAMILY_NAME"
	
	if Series<>"" then
		SQL=SQL+" ,SN.SERIES_NAME"
	end if
	
	if SubSeries<>"" then
		SQL=SQL+" ,SS.SUBSERIES_NAME"
	end if
	
	if Model<>"" then
		SQL=SQL+" ,TEMP.part_number_tag"
	end if
	SQL=SQL+" ,TEMP.defect_name"
	SQL=SQL+" order BY FAMILY_NAME"
	if Series<>"" then
		SQL=SQL+" ,SN.SERIES_NAME"
	end if
	
	if SubSeries<>"" then
		SQL=SQL+" ,SS.SUBSERIES_NAME"
	end if
	
	if Model<>"" then
		SQL=SQL+" ,TEMP.part_number_tag"
	end if
	
	SQL=SQL+",sum(temp.defect_quantity) desc "
	
	RESPONSE.WRITE SQL
	response.end
	rs.open SQL,conn,1,3
				
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	function LoadNextItem()
	{
		form1.action="DefectReport.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="DefectReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>
<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="18"> </td>
  </tr>
  <tr>
	<%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
    <td height="20" class="t-t-Borrow"><div align="center"><%=rs.Fields(i).name%> </div></td>
  	<%
			next 
		end if
	%>
</tr>
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	<%			
				for i=0 to rs.Fields.count-1

	%>
				 <td height="20" ><div align="center"><%=rs(i).value%> </div></td>
				 
  	<%			
				next
				rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
 
</table>
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->