<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/HVOEM_Open.asp" -->
<!--#include virtual="/WOCF/PVS_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<% Server.ScriptTimeout=999 %>
<script type="text/javascript">
	window.opener=null;   
  	window.open("","_self");  
	window.close();
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

<body bgcolor="#339966">
<%

sql="Delete from HVreport"
conn.execute(sql)

function TODB (TODBvalue)
	
	if  isnull(TODBvalue) then
         TODB=""
	else
		TODB=cstr(TODBvalue)
	end if
	
	
end function

sql="SELECT  *  FROM [DWH].[dbo].[ReportDataDaily]"
'"
rsHVOEM.open sql,connHVOEM,1,3
	
i=0
while not rsHVOEM.eof
	PRODUCT=trim(TODB(rsHVOEM("PRODUCT")))
	LINE_NUM=trim(TODB(rsHVOEM("LINE_NUM")))
	CREATE_TIME=trim(cstr(rsHVOEM("CREATE_TIME")))


	LINE_TYPE=TODB(rsHVOEM("LINE_TYPE"))
	TARGET=TODB(rsHVOEM("TARGET"))
	ACTUAL=TODB(rsHVOEM("ACTUAL"))
	REAL_TO_WH=TODB(rsHVOEM("REAL_TO_WH"))
	PQ_OUT=TODB(rsHVOEM("PQ_OUT"))
	PQ_START=TODB(rsHVOEM("PQ_START"))
	PQ_TARGET=TODB(rsHVOEM("PQ_TARGET"))
	PQ_YIELD=TODB(rsHVOEM("PQ_YIELD"))
	AMS_GOOD=TODB(rsHVOEM("AMS_GOOD"))
	AMS_START=TODB(rsHVOEM("AMS_START"))
	AMS_TARGET=TODB(rsHVOEM("AMS_TARGET"))
	AMS_YIELD=TODB(rsHVOEM("AMS_YIELD"))
	COSMETIC_GOOD=TODB(rsHVOEM("COSMETIC_GOOD"))
	COSMETIC_START=TODB(rsHVOEM("COSMETIC_START"))
	COSMETIC_TARGET=TODB(rsHVOEM("COSMETIC_TARGET"))
	COSMETIC_YIELD=TODB(rsHVOEM("COSMETIC_YIELD"))
	TTY_OPEN=TODB(rsHVOEM("TTY_OPEN"))
	TTY_CLOSE=TODB(rsHVOEM("TTY_CLOSE"))
	TTY_TARGET=TODB(rsHVOEM("TTY_TARGET"))
	TTY_YIELD=TODB(rsHVOEM("TTY_YIELD"))
	'sqlP="select * from pvs_tobps where ad_id='"+ad_id+"'"
	'rs.open sqlP,conn,1,3
	'if rs.bof and rs.eof then
	'sql="insert into HvReport (PRODUCT,LINE_NUM,CREATE_TIME,LINE_TYPE,TARGET,ACTUAL,REAL_TO_WH,PQ_OUT,PQ_START,PQ_TARGET,PQ_YIELD,AMS_GOOD,AMS_START,AMS_TARGET,AMS_YIELD,COSMETIC_GOOD,COSMETIC_START,COSMETIC_TARGET,COSMETIC_YIELD,TTY_OPEN,TTY_CLOSE,FPY_Target,FPY) VALUES  ('"+PRODUCT+"','"+LINE_NUM+"','"+CREATE_TIME+"','"+LINE_TYPE+"','"+TARGET+"','"+ACTUAL+"','"+REAL_TO_WH+"','"+PQ_OUT+"','"++"','"+PQ_TARGET+"','"+PQ_YIELD+"','"+AMS_GOOD+"','"+AMS_START+"','"+AMS_TARGET+"','"+AMS_YIELD+"','"+COSMETIC_GOOD+"','"+COSMETIC_START+"','"+COSMETIC_TARGET+"','"+COSMETIC_YIELD+"','"+TTY_OPEN+"','"+TTY_CLOSE+"','"+FPY_Target+"','"+FPY+"')"
	
	 sql="insert into HvReport (PRODUCT,LINE_NUM,CREATE_TIME,LINE_TYPE,TARGET,ACTUAL,REAL_TO_WH,PQ_OUT,PQ_START,PQ_TARGET,PQ_YIELD,AMS_GOOD,AMS_START,AMS_TARGET,AMS_YIELD,COSMETIC_GOOD,COSMETIC_START,COSMETIC_TARGET,COSMETIC_YIELD,TTY_OPEN,TTY_CLOSE,TTY_TARGET,TTY_YIELD) VALUES  ('"&PRODUCT&"','"&LINE_NUM&"','"&CREATE_TIME&"','"&LINE_TYPE&"','"&TARGET&"','"&ACTUAL&"','"&REAL_TO_WH&"','"&PQ_OUT&"','"&PQ_START&"','"&PQ_TARGET&"','"&PQ_YIELD&"','"&AMS_GOOD&"','"&AMS_START&"','"&AMS_TARGET&"','"&AMS_YIELD&"','"&COSMETIC_GOOD&"','"&COSMETIC_START&"','"&COSMETIC_TARGET&"','"&COSMETIC_YIELD&"','"&TTY_OPEN&"','"&TTY_CLOSE&"','"&TTY_TARGET&"','"&TTY_YIELD&"')"
	 
	 response.Write(sql)
	 	response.Write("<BR>")
'sql="insert into HvReport (PRODUCT,LINE_NUM,LINE_TYPE) VALUES  ('RA','4','auto')"
	 conn.execute(sql)
	 i=i+1
    'end if
	'rs.close
rsHVOEM.movenext		
wend
rsHVOEM.close







L=i+K




response.Write(L)
'response.End()
endtime=timer()
COST_TIME=cstr(FormatNumber((endtime-startime),2,-1))&"秒"

sqlL="insert into HVOEM_LOG (HVOEM_COUNT,COST_TIME,crt_time) VALUES  ('"+cstr(L)+"','"+COST_TIME+"','"+cstr(now())+"')"

conn.execute(sqlL)
%>
<p>
页面执行时间：<font size="+1" color="#0000FF"><b><%=COST_TIME%></b></font></p>

</body>
</html>
<!--#include virtual="/WOCF/HVOEM_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->