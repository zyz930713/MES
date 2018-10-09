<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<%
filePath=server.mappath("\Devices\Excel")

Set obj =Server.CreateObject("LyfUpload.UploadFile")
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

SQL="delete from BARCODE_SCANNER"
rs.open SQL,conn,3,3
SQL="select * from BARCODE_SCANNER"
rs.open SQL,conn,3,3
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\imtemp.xls")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

i = 2
Do
rs.AddNew
rs("NID")="SA"&NID_SEQ("SCANNER")
rs("SCANNER_NAME") = trim(Sheet.Cells(i,1).Text)
rs("BRADE_ADDRESS") = trim(Sheet.Cells(i,2).Text)
rs("SCANNER_ADDRESS") = trim(Sheet.Cells(i,3).Text)
rs("LINK_COMPUTER_NAME") = trim(Sheet.Cells(i,4).Text)
rs.Update
i = i + 1
Loop Until Trim(Sheet.Cells(i, 1).Text) = "end" or Trim(Sheet.Cells(i, 1).Text)=""

rs.close
OExcel.workbooks.close
OExcel.quit 
set OExcel=Nothing

DeleteFile filePath&"\*.xls"
%>
<script language="JavaScript">
alert("Successfully upload scanners")
window.close()
</script>
<%
end if
set obj=nothing
%>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->