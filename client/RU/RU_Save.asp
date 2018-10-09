<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

Action=request("Action")
JOB_number=request("JOB_number")
response.Write(action)
response.Write("<BR>")
response.Write(JOB_number)

FINAL_SCRAP_QUANTITY=trim(request("FINAL_SCRAP_QUANTITY"))
REWORK_GOOD_QUANTITY=trim(request("REWORK_GOOD_QUANTITY"))
response.Write("<BR>")
response.Write(FINAL_SCRAP_QUANTITY)
if Action="EDIT" and JOB_number <>"" then
conn.execute("update JOB_master set FINAL_SCRAP_QUANTITY='"&FINAL_SCRAP_QUANTITY&"' ,REWORK_GOOD_QUANTITY='"&REWORK_GOOD_QUANTITY&"' where job_number='"&JOB_number&"' ")
response.Redirect ("RU1.asp")
end if
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->