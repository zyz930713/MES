<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetMachineStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
station=request.QueryString("station")
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
	stationstartquantity=0
	stationtotaldefectcodequantity=0
	station_assembly_yield=0
	txtDF="&nbsp;"
	SQL="select STATION_START_QUANTITY,STATION_DEFECTCODE_QUANTITY,STATION_ASSEMBLY_YIELD from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and STATION_ID='"&station&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		stationstartquantity=getJobActionValue(jobnumber,sheetnumber,station,"AC00000097")
		txtDF=GetMachineStationDefectCode(jobnumber,sheetnumber,station,stationtotaldefectcodequantity)
		station_assembly_yield=csng(stationstartquantity-stationtotaldefectcodequantity)/csng(stationstartquantity)
		rs("STATION_START_QUANTITY")=stationstartquantity
		rs("STATION_DEFECTCODE_QUANTITY")=stationtotaldefectcodequantity
		rs("STATION_ASSEMBLY_YIELD")=station_assembly_yield
		rs.update
	end if
	rs.close
set rs=nothing%>
<script language="javascript">
parent.document.all.y<%=jobnumber%>_<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%>_<%=station%>.innerHTML="<table border='0' align='center' cellpadding='0' cellspacing='0'><tr><td><%=txtDF%></td></tr><tr><td><div align='center'><%=formatpercent(station_assembly_yield,2,-1)%><%if station_assembly_yield<>0 then%><img src='/Images/Refresh1.gif' alt='Click button to refresh value' width='15' height='15' align='absmiddle' style='cursor:hand' onClick=javascript:document.all.RefreshFrame.src='StationYield_Refresh.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&station=<%=station%>'><%end if%></div></td></tr></table>"
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->