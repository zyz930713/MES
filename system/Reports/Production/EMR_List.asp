<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>


<!--#include virtual="/WOCF/BOCF_Open.asp" -->



<%

job_number=request("job_number")

subjobnumber=request("sheet_number")
sql="select * from material_count_record  M , mr_dispatch MR where m.labelid = mr.label_no and m.job_number='"&job_number&"' and m.subjobnumber='"&subjobnumber&"'"

rs.open SQL,conn,1,3

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->

<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>

<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<style type="text/css">
.t-t-DarkBlue div {
	font-size: 18px;
}
</style>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Packing_Plan/addPacking_Plan.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>

<tr>
<td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">JOB Number</div></td>
   <td height="20" class="t-t-Borrow"><div align="center">Sheet Number</div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><SPAN space="" right-pos="0|6" left-pos="0|6" tabcount="-1">Label</SPAN> NO</div></td>
   <td height="20" class="t-t-Borrow"><div align="center">12NC</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Lot NO</div></td>
 
  

  
  

  </tr>
<%
i=1
if not rs.eof then

while not rs.eof  
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
 
    
	<td height="20"><div align="center"><%= rs("job_number") %></div></td>
	<td><%= rs("subjobnumber")%>&nbsp;</td>
	<td><%= rs("LABELID")%>&nbsp;</td>
	<td><%= rs("material_part_number")%>&nbsp;</td>
	<%if rs("material_part_number")=430307703141 OR rs("material_part_number")=430307703241     OR  rs("material_part_number")=430307704631  then%>
    <td><a href="Pot_list.asp?material_part_number=<%=rs("material_part_number")%>&material_lot_number=<%=rs("material_lot_number")%>">	<%= rs("material_lot_number")%></a>&nbsp;</td>	
    <%else%>
    
    <td><%= rs("material_lot_number")%>&nbsp;</td>	
    <%end if%>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="13" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>

</table>

<%

partnumber=request("partnumber")
set rsM=server.createobject("adodb.recordset")
if partnumber="240326300129" then

	sqlM="select * from job_actions where job_number='"&job_number&"' and sheet_number='"&subjobnumber&"' and  action_id='AC00044009'"
	rsM.open sqlM,conn,1,3
		if not rsM.eof then
		ACTION_VALUE=rsM("ACTION_VALUE")
		end if 
	rsM.close

	tArr = Split(ACTION_VALUE,",")
	 For i = 0 To UBound(tArr)  ' UBound(tArr) 遍历数组有多少个
			NDcode=(tArr(i))
			ajobnumber=split(NDcode,"-")
			'
			jobnumber=left(NDcode,len(NDcode)-4)
			
			jobnumberALL=jobnumberALL&","&jobnumber
			
			strSheetNo=ajobnumber(ubound(ajobnumber))
			
			if isnumeric(strSheetNo) then
				sheetnumber = cint(strSheetNo)
				sheetnumberALL=sheetnumberALL&","&sheetnumber
				sqlM="select * from material_count_record  WHERE JOB_NUMBER='"&jobnumber&"' and SUBJOBNUMBER='"&sheetnumber&"'"
				rsM.open sqlM,conn,1,3
					if not rsM.eof then
					LABELIDN=LABELIDN&rsM("LABELID")&","
					end if	
				rsM.close
			end if
			
	 Next

	

	if LABELIDN<>"" then
		LABELIDN=left(LABELIDN,len(LABELIDN)-1)
	end if
'select * from material_count_record  WHERE JOB_NUMBER='KB85111350E' and SUBJOBNUMBER='2'
else

	sqlM="select * from material_count_record  where job_number='"&job_number&"' and subjobnumber='"&subjobnumber&"' and substr(labelid,0,2)<>'LB'"
	'response.Write(sqlM)
	rsM.open sqlM,conn,1,3
	if not rsM.eof then
	LABELIDN=rsM("LABELID")
	end if 
	rsM.close

end if

set rsE=server.createobject("adodb.recordset")
'sqlE="select * from emr_pack_detail where box_id='"&LABELIDN&"'"
sqlE=" select * from emr_pack_detail where box_id  in('"&replace(LABELIDN,",","','")&"')"

rsE.open sqlE,conn,1,3
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">

<TR><td colspan="20" class="t-t-DarkBlue"><div align="center">MEMBRANE ASSY</div></td></TR>
<tr align="center">
    <td height="20" class="t-t-Borrow"><div align="center">BOX ID</div></td>
    <td class="t-t-Borrow">12NC</td>
    <td height="20" class="t-t-Borrow"><div align="center">Production</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Shift</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Packing Time</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Packing QTY</div></td>
    <td class="t-t-Borrow">PLATE</td>
    <td class="t-t-Borrow">FOIL</td>
    <td class="t-t-Borrow">RING</td>
    <td class="t-t-Borrow">Equipment NO</td>
    <td class="t-t-Borrow">Inspector</td>
    <td class="t-t-Borrow">Form</td>
    <td class="t-t-Borrow">Deep drawing</td>
</tr>
<%
i=1
if not rsE.eof then

while not rsE.eof  
%>
<tr  align="center">
    <td height="20"><%= rsE("BOX_ID") %></td>
    <td height="20"><%= rsE("PART_NUMBER") %></td>
    <td><%= rsE("PRODUCTION_NAME")%></td>
    <td><%= rsE("SHIFT")%></td>
    <td><%= rsE("PACK_TIME")%></td>
    <td><%= rsE("PACK_QTY")%></td>
    <td><%= rsE("SUPPLIERA")%></td>
    <td><%= rsE("SUPPLIERB")%></td>
    <td><%= rsE("SUPPLIERC")%></td>
    <td><%= rsE("EQUIPMENT_NO")%></td>
    <td><%= rsE("INSPECTOR")%></td>
    <td><%= rsE("FORMA")%>&nbsp;&nbsp;<%= rsE("FORMB")%>&nbsp;&nbsp;<%= rsE("FORMC")%>&nbsp;&nbsp;<%= rsE("FORMD")%></td>
    <td><%= rsE("DEEPA")%>&nbsp;&nbsp;<%= rsE("DEEPB")%>&nbsp;&nbsp;<%= rsE("DEEPC")%>&nbsp;&nbsp;<%= rsE("DEEPD")%></td>	
</tr>
<%
i=i+1
rsE.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="20" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rsE.close%>

</table>

<%
if  jobnumberALL<>"" and sheetnumberALL<>"" then

sqlE="select a.job_number, a.sheet_number,b.station_description, a.batch_no, a.product_time, a.scan_datetime from MATERIAL_COUNT_RECORD_NEW a, MATERIAL_STATION b where a.sms_id= b.station_no  and a.job_number in('"&replace(jobnumberALL,",","','")&"') and a.sheet_number in('"&replace(sheetnumberALL,",","','")&"') order by a.job_number,a.sheet_number desc"
else
sqlE="select a.job_number, a.sheet_number,b.station_description, a.batch_no, a.product_time, a.scan_datetime from MATERIAL_COUNT_RECORD_NEW a, MATERIAL_STATION b where a.sms_id= b.station_no  and a.job_number ='"&job_number&"' and a.sheet_number ='"&subjobnumber&"' order by a.job_number,a.sheet_number desc"
end if

rsE.open sqlE,conn,1,3%>


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">

<TR><td colspan="12" class="t-t-DarkBlue"><div align="center">Glue&amp;Copper</div></td></TR>
<tr align="center">
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Station</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Batch NO</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Product Time</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Scan Time</div></td>
  </tr>
<%
i=1
if not rsE.eof then

while not rsE.eof  
%>
<tr  align="center">
    <td height="20"><%=i%></td>
    <td><%=rsE("station_description")%></td>
    <td><%=rsE("batch_no")%></td>
    <td><%=rsE("product_time")%></td>
    <td><%=rsE("scan_datetime")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
<%
i=i+1
rsE.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="12" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rsE.close%>

</table>
<table width="1154" height="26">
<tr><td align="center"><input name="btnClose2" type="button"  id="btnClose2"  value="Back" onClick="javascript:history.go(-1);"></td>
</tr></table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->

