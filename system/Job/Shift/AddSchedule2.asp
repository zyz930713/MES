<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%

path=server.MapPath("\Job\Shift\Bat")
filename="\Test.bat"
runtime="200610280911"
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(path&filename, True)
a.WriteLine("Ping KES-WEB -t")
a.Close
set a =nothing
set fs=nothing

strComputer = "."
Set objService = GetObject("winmgmts:\\" & strComputer)
Set objNewJob = objService.Get("Win32_ScheduledJob")
errJobCreated = objNewJob.Create (path&filename,runtime&"00.000000+480",false,,,true,JobID)
If Err.Number = 0 Then
 word="A New Schedule ("&JobID&") is created and run at "&runtime
Else
 word( "An error occurred: " & errJobCreated)
End If
set objNewJob=nothing
set objService=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>
