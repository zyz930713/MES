<%
if session("JOB_NUMBER")="" or session("SHEET_NUMBER")="" or session("JOB_TYPE")="" then
response.Redirect("/KES1/Station1.asp?errorstring=Idle time is too long.<br>空闲时间超出！")
end if
if session("stationright")=false then
response.Redirect("/KES1/Station1.asp?errorstring=You have not right to transact this staion.<br>您无权处理本站！")
end if
%>