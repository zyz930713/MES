<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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

dim startime
startime=timer()
fromTime=now()
response.Write(fromTime)
response.Write("<BR>")
toTime=DateAdd("n",-60,fromTime)
response.Write(toTime)
response.Write("<BR>")
Fe=Minute(fromTime)
Se=Second(fromTime)
response.Write(Fe&"分")
response.Write("<BR>")
response.Write(Se&"秒")
response.Write("<BR>")
		if Fe>=0 and Fe<30 then
		toTimeF=DateAdd("n",-Fe,toTime)
		elseif  fe>=30 and fe<60 then
		toTimeF=DateAdd("n",-(Fe-30),toTime)
		end if
'ho=hour(toTimeS)
'response.Write(ho)
'response.Write("<BR>")

toTimeS=DateAdd("s",-Se,toTimeF)

response.Write(toTimeS)
response.Write("<BR>")
toTimeEnd=dateadd("s",59,dateadd("n",29,toTimeS))
response.Write(toTimeend)
response.Write("<BR>")






'response.End()

function TODB (TODBvalue)
	
	if  isnull(TODBvalue) then
         TODB=""
	else
		TODB=cstr(TODBvalue)
	end if
	
	
end function

sql="SELECT  ad_id ,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,measurecount,product_id,serialnumber  FROM [pvs].[dbo].[vw_adid_by_sn_pvs]  where  measuredatetime BETWEEN  '"&toTimeS&"'  and    '"&toTimeend&"'"
'sql="SELECT  ad_id ,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,serialnumber  FROM [pvs].[dbo].[vw_adid_by_sn_pvs]  where  measuredatetime BETWEEN  '"&toTimeS&"'  and    '2013-12-26 12:30:01.000'"
rsPVS.open sql,connPVS,1,3
	
i=0
while not rsPVS.eof
	ad_id=cstr(rsPVS("ad_id"))
	line_id=cstr(rsPVS("line_id"))
	measuredatetime=cstr(rsPVS("measuredatetime"))
	AMS_line_id=TODB(rsPVS("AMS_line_id"))
	
	AMS_measuredatetime=TODB(rsPVS("AMS_measuredatetime"))
	
	adfail=cstr(rsPVS("adfail"))
	aderror_id=TODB(rsPVS("aderror_id"))
	adcriterionerror_id=TODB(rsPVS("adcriterionerror_id"))
	measurementpc_id=cstr(rsPVS("measurementpc_id"))
	measurecount=TODB(rsPVS("measurecount"))
	product_id=TODB(rsPVS("product_id"))
	serialnumber=cstr(rsPVS("serialnumber"))
	'sqlP="select * from pvs_tobps where ad_id='"+ad_id+"'"
	'rs.open sqlP,conn,1,3
	'if rs.bof and rs.eof then
	 sql="insert into PVS_TOBPS (ad_id,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,measurecount,product_id,serialnumber) VALUES  ('"+ad_id+"','"+line_id+"','"+measuredatetime+"','"+AMS_line_id+"','"+AMS_measuredatetime+"','"+adfail+"','"+aderror_id+"','"+adcriterionerror_id+"','"+measurementpc_id+"','"+measurecount+"','"+product_id+"','"+serialnumber+"')"
	 conn.execute(sql)
	 i=i+1
    'end if
	'rs.close
rsPVS.movenext		
wend
rspvs.close



sql="SELECT  ad_id ,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,measurecount,product_id,serialnumber  FROM [pvs].[dbo].[vw_adid_by_sn_pvs]  where  AMS_measuredatetime BETWEEN  '"&toTimeS&"'  and    '"&toTimeend&"'"
'sql="SELECT  ad_id ,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,serialnumber  FROM [pvs].[dbo].[vw_adid_by_sn_pvs]  where  measuredatetime BETWEEN  '"&toTimeS&"'  and    '2013-12-26 12:30:01.000'"
rsPVS.open sql,connPVS,1,3
	
K=0
while not rsPVS.eof
	ad_id=cstr(rsPVS("ad_id"))
	line_id=cstr(rsPVS("line_id"))
	measuredatetime=cstr(rsPVS("measuredatetime"))
	AMS_line_id=TODB(rsPVS("AMS_line_id"))
	
	AMS_measuredatetime=TODB(rsPVS("AMS_measuredatetime"))
	
	adfail=cstr(rsPVS("adfail"))
	aderror_id=TODB(rsPVS("aderror_id"))
	adcriterionerror_id=TODB(rsPVS("adcriterionerror_id"))
	measurementpc_id=cstr(rsPVS("measurementpc_id"))
	measurecount=TODB(rsPVS("measurecount"))
	product_id=TODB(rsPVS("product_id"))
	
	serialnumber=cstr(rsPVS("serialnumber"))
	'sqlP="select * from pvs_tobps where ad_id='"+ad_id+"'"
	'rs.open sqlP,conn,1,3
	'if rs.bof and rs.eof then
	 sql="insert into PVS_TOBPS (ad_id,line_id,measuredatetime,AMS_line_id,AMS_measuredatetime,adfail,aderror_id,adcriterionerror_id,measurementpc_id,measurecount,product_id,serialnumber) VALUES  ('"+ad_id+"','"+line_id+"','"+measuredatetime+"','"+AMS_line_id+"','"+AMS_measuredatetime+"','"+adfail+"','"+aderror_id+"','"+adcriterionerror_id+"','"+measurementpc_id+"','"+measurecount+"','"+product_id+"','"+serialnumber+"')"
	 conn.execute(sql)
	 K=K+1
    'end if
	'rs.close
rsPVS.movenext		
wend





L=i+K




response.Write(L)
'response.End()
endtime=timer()
COST_TIME=cstr(FormatNumber((endtime-startime),2,-1))&"秒"

sqlL="insert into PVS_LOG (PVS_COUNT,START_TIME,END_TIME,COST_TIME,crt_time) VALUES  ('"+cstr(L)+"','"+cstr(toTimeS)+"','"+cstr(toTimeend)+"','"+COST_TIME+"','"+cstr(fromTime)+"')"

conn.execute(sqlL)
%>
<p>
页面执行时间：<font size="+1" color="#0000FF"><b><%=COST_TIME%></b></font></p>

</body>
</html>
<!--#include virtual="/WOCF/PVS_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->