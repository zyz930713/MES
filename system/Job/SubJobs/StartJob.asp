<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/HoldReleaseFunc.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
reason=request.QueryString("reason")

response.Redirect("BatchRelease.asp?Action=Query&jobnumber="&jobnumber&"&sheetnumber="&sheetnumber)

SQL="select STATUS,CURRENT_STATION_ID from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
current_station_id=rs("CURRENT_STATION_ID")
rs("STATUS")=0
rs.update
rs.close
SQL="select CONTROL_TYPE,CONTROL_STATION,CONTROL_REASON,CONTROL_PERSON,CONTROL_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
rs("CONTROL_TYPE")=rs("CONTROL_TYPE")&"0"&","
rs("CONTROL_STATION")=rs("CONTROL_STATION")&current_station_id&","
rs("CONTROL_REASON")=rs("CONTROL_REASON")&reason&","
rs("CONTROL_PERSON")=rs("CONTROL_PERSON")&session("code")&","
rs("CONTROL_TIME")=rs("CONTROL_TIME")&now()&","
rs.update
rs.close

'new method
SQL="SELECT MAX(ID) FROM JOB_HOLD_RELEASE_HISTORY WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+sheetnumber+"'"
SQL=SQL+" AND STATION_ID='"+current_station_id+"' AND TRANSACTION_TYPE ='1'"
set rsMaxHoldID=server.createobject("adodb.recordset")
rsMaxHoldID.open SQL,conn,1,3
MaxHoldID=""
if rsMaxHoldID.recordcount>0 then 
	MaxHoldID=rsMaxHoldID(0)
end if 

SQL="insert into JOB_HOLD_RELEASE_HISTORY (JOB_NUMBER,SHEET_NUMBER,STATION_ID,TRANSACTION_TYPE,TRANSACTION_PERSON,TRANSACTION_TIME,TRANSACTION_REASON,ID,RELATED_ID)"
SQL=SQL+"VALUES('"+jobnumber+"','"+sheetnumber+"','"+current_station_id+"','0','"+session("code")+"',"
SQL=SQL+" '"+CSTR(now())+"','"+reason+"','"+cstr(NID_SEQ("TRANSACTION"))+"','"+CSTR(MaxHoldID)+"')"
set rsHoldJob2=server.createobject("adodb.recordset")
rsHoldJob2.open SQL,conn,1,3

'add by Lennie Wei 2013-01-10  
jobNum = jobnumber & "-" & string(3-len(sheetnumber),"0")&sheetnumber   
sendHoldReleaseNotiInfo jobNum,"Release",reason
'end add by Lennie Wei 2013-01-10

%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>