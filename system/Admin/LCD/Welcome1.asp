<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Upload_Start.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=Upload.Form("path")
query=Upload.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query

Set File = Upload.Files("attachment")
If Not File Is Nothing Then
	file_name=File.FileName
	file_extension=File.Ext
	file_size=File.Size
	if (file_extension=".ppsx" or file_extension=".pps") and file_size<=1024*1024*5 then
		thisday=Upload.Form("thisday")
		SQL="select * from DAILY_WELCOME where to_char(WELCOME_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		rs("WELCOME_DAY")=thisday
		rs("WELCOME_TYPE")=2
		end if
		rs("FILE_NAME")=file_name
		rs("CREATOR_CODE")=session("code")
		rs.update
		rs.close
		file.SaveAs application("LCD_Welcome")&"\"&file_name '保存文件到服务器
		word="Successfully upload file."
		action="location.href='"&beforepath&"'"
	else
	word="File format is not pps or size exceeds 5 MB."
	action="history.back()"
	end if
else
	word="File is wrong."
	action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/Functions/Upload_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->