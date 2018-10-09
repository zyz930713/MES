<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<%
Set obj =Server.CreateObject("LyfUpload.UploadFile")
schedule_id=obj.request("schedule_id")
year_index=obj.request("yearindex")
week_index=obj.request("weekindex")
week_start_day=obj.request("weekstartday")
week_end_day=obj.request("weekendday")
schedule_name=obj.request("schedulename")
re_upload=obj.request("loadagain")
path=obj.request("path")
query=obj.request("query")
query=replace(query,"*","&")
beforepath=path&"?"&query

SQL="select * from JOB_SCHEDULE_LIST where NID='"&schedule_id&"'"
rs.open SQL,conn,3,3
' 1 if
if not rs.eof then
rs("SCHEDULE_NAME")=schedule_name
rs("YEAR_INDEX")=year_index
rs("WEEK_INDEX")=week_index
rs("WEEK_START_DAY")=week_start_day
rs("WEEK_END_DAY")=week_end_day
rs("CREATOR_CODE")=session("code")
rs("UPDATE_TIME")=now()
rs.update
' 1 if
end if
rs.close

'2 if 
if re_upload="1" then
filePath=server.mappath("\Reports\Excel")
obj.maxsize=1047961600  '设置文件上传的最大为1024*1024*10个字节(10M)
obj.extname="xls"
ss=obj.SaveFile("file", filePath, true,"imtemp.xls")
aa=obj.filetype("file") '得到文件的Content-Type

if ss= "" then
%>
<script language="JavaScript">
alert("Fail to upload file!")
window.close()
</script>
<%
elseif ss= "0"  then
%>
<script language="JavaScript">
alert("File size is too large!")
window.close()
</script>
<%
elseif ss= "1" then
%>
<script language="JavaScript">
alert("Uploaded file is not excel file!")
window.close()
</script>
<%
else

SQL="delete from JOB_SCHEDULE_DETAIL where SCHEDULE_ID='"&schedule_id&"'"
rs.open SQL,conn,3,3
SQL="select * from JOB_SCHEDULE_DETAIL"
rs.open SQL,conn,3,3
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\imtemp.xls")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

i = 2
Do
rs.AddNew
rs("SCHEDULE_ID")=schedule_id
rs("JOB_NUMBER") = trim(Sheet.Cells(i,1).Text)
rs("PART_NUMBER_ID") = getPart(false,"JOB_SCHEDULE",""," where P.PART_NUMBER='"&trim(Sheet.Cells(i,2).Text)&"'","","","","","")
rs("QUANTITY") = trim(Sheet.Cells(i,3).Text)
rs("START_TIME") = Sheet.Cells(i,4).Text
rs("COMPLETE_TIME") = Sheet.Cells(i,5).Text
rs.Update
i = i + 1
Loop Until Trim(Sheet.Cells(i, 1).Text) = "end" or Trim(Sheet.Cells(i, 1).Text)=""

rs.close
OExcel.workbooks.close
OExcel.quit 
set OExcel=Nothing

set myFs=createObject("scripting.FileSystemObject") 
myFs.DeleteFile filePath&"\*.xls"
set myFs=nothing
end if
%>
<script language="JavaScript">
alert("Successfully re-upload Job Schedule of <%=year_index%>-<%=week_index%>")
location.href="<%=beforepath%>"
</script>
<%
else
%>
<script language="JavaScript">
alert("Successfully edit Job Schedule of <%=year_index%>-<%=week_index%>")
location.href="<%=beforepath%>"
</script>
<%end if
set obj=nothing%>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->