<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
actionString=Request.QueryString("actionString")
defectString=Request.QueryString("defectString")
typeString=Request.QueryString("typeString")
categoryString=Request.QueryString("categoryString")

GuId=request.Form("txtGuid")
StationId=request.Form("StationId")
RoutingId=request.Form("RoutingId")
MeterialCount=request.Form("txtCount")

'delete temp
SQL_DA=" delete from ROUTING_ACTION_DETAIL_TEMP where GUID='"&GuId&"' and STATION_ID='"&StationId&"'"
rs.open SQL_DA,conn,1,3
SQL_DD=" delete from ROUTING_DEFECT_DETAIL_TEMP where GUID='"&GuId&"' and STATION_ID='"&StationId&"'"
rs.open SQL_DD,conn,1,3

'add new data to action and defectcode
SQL="select * from ROUTING_ACTION_DETAIL_TEMP where GUID='"&GuId&"' and STATION_ID='"&StationId&"'"
rs.open SQL,conn,1,3
if not  rs.eof then
	word="Action has existed, please input again."
	action="history.back()"
else
	if actionString <>"" then
		Dim action 
		action=Split(actionString,",")
		For   i   =   Lbound(action)   to   Ubound(action) 			 
			rs.addnew
			rs("GUID")=GuId
			rs("STATION_ID")=StationId
			rs("STATION_SEQENCE")="0"
			rs("ACTION_ID")=action(i)
			rs("ACTION_SEQENCE")=i+1
			rs.update
			'rs.close

		Next 
	end if
end if	

SQL="select * from ROUTING_DEFECT_DETAIL_TEMP where GUID='"&GuId&"' and STATION_ID='"&StationId&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open SQL,conn,1,3
if not  rs1.eof then
	word="DefectCode has existed, please input again."
	action="history.back()"
else
	if defectString<>"" then
		Dim DefectCode,types
		DefectCode=Split(defectString,",")
		types=Split(typeString,",")
		For   j   =   Lbound(DefectCode)   to   Ubound(DefectCode) 	
			rs1.addnew
			rs1("GUID")=GuId
			rs1("STATION_ID")=StationId
			rs1("STATION_SEQENCE")="0"
			rs1("DEFECTCODE_ID")=DefectCode(j)
			rs1("DEFECTCODE_SEQENCE")=j+1
			rs1("DEFECTCODE_TRANSACTION_TYPE")=types(j)
			rs1.update
			
		Next 
	end if	
end if

'save Meterial Count
SQL="Select * from MATERIAL_COUNT where ROUTING_ID='"&RoutingId&"' and STATION_ID='"&StationId&"'"
set rsM=server.createobject("adodb.recordset")
rsM.open SQL,conn,1,3
if rsM.recordcount=0 then 'add new
	rsM.addnew
	rsM("ROUTING_ID")=RoutingId
	rsM("STATION_ID")=StationId
	rsM("MATERIAL_COUNT")=MeterialCount
	rsM.update
else
	rsM("ROUTING_ID")=RoutingId
	rsM("STATION_ID")=StationId
	rsM("MATERIAL_COUNT")=MeterialCount
	rsM.update
end if

'add by jack zhang 2012-4-10
SQL="DELETE from STATION_MATERIAL_BOM where ROUTING_ID='"&RoutingId&"' and STATION_ID='"&StationId&"'"
set rsBOMDELETE=server.createobject("adodb.recordset")
rsBOMDELETE.open SQL,conn,1,3

SQL="select * from STATION_MATERIAL_BOM where ROUTING_ID='"&RoutingId&"' and STATION_ID='"&StationId&"'"
set rsBOM=server.createobject("adodb.recordset")
rsBOM.open SQL,conn,1,3
if categoryString<>"" then
	categoryidarr=Split(categoryString,",")
	For   j   =   Lbound(categoryidarr)   to   Ubound(categoryidarr) 	
	
		SQL=" select b.station_chinese_name, c.category_name from STATION_MATERIAL_BOM a, station_new b, material_category c"
		SQL=SQL+" where a.station_id=b.nid and a.category_id= c.category_id and a.ROUTING_ID='"&RoutingId&"' and a.CATEGORY_ID='"&categoryidarr(j)&"'"
		set rsExisting=server.createobject("adodb.recordset")
		rsExisting.open SQL,conn,1,3
		if(rsExisting.recordcount>0) then
				word="Category:"+rsExisting("category_name")+" 已经在"+rsExisting("station_chinese_name")+"维护了!"
				word=replace(word,"'","")
				response.Redirect "SetActionDefectCode.asp?word="+word+"&StationId="+StationId+"&GUID="+GuId+"&rid="+RoutingId

		else
			rsBOM.addnew
			rsBOM("ROUTING_ID")=RoutingId
			rsBOM("STATION_ID")=StationId
			rsBOM("CATEGORY_ID")=categoryidarr(j)
			rsBOM("SEQUENCE_ID")=j+1
			rsBOM.update
		end if 
	Next 
end if	



word="Set Successfully!"
url="window.close()"
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
	alert("<%=word%>");
	<%=url%>
</script>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->