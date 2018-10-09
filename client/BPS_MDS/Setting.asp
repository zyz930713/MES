<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Set</title>
<script language="javascript" type="text/javascript" src="include/ThreeMenu.js"></script>
<style type="text/css">

</style>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<%
	ProductName = request("ProductName")
	LineName = request("LineName")
	LinePC = request("LinePC")
	DDN = request("DDN")
	HShift = request("HShift")
	Ddate = request("Ddate")

	
	'取Cookies缓存数据'
	if ProductName = "" then
		ProductName = request.Cookies("mds")("ProductName")
	else
		response.Cookies("mds")("ProductName") = ProductName
	end if
	
	if LineName = "" then
		LineName = request.Cookies("mds")("LineName")
	else
		response.Cookies("mds")("LineName") = LineName
	end if
	
	if LinePC = "" then
		LinePC = request.Cookies("mds")("LinePC")
	else
		response.Cookies("mds")("LinePC") = LinePC
	end if
	'取Cookies缓存数据'
	
	response.write LinePC
%>
<center>
<div >
	<form name="form1" action="Setting.asp" method="post">
		线号：<select name="LineName">
				<option value="AF01" <%if LineName = "AF01" then response.write "selected"%> >AF01</option>
				<option value="AF02" <%if LineName = "AF02" then response.write "selected"%> >AF02</option>
				<option value="AF03" <%if LineName = "AF03" then response.write "selected"%> >AF03</option>
				<option value="AF04" <%if LineName = "AF04" then response.write "selected"%> >AF04</option>
				<option value="AF05" <%if LineName = "AF05" then response.write "selected"%> >AF05</option>
			</select>

		是否属于产线电脑：<select name="LinePC">
			<option value="1" <%if LinePC = "1" then response.write "selected"%> >Yes</option>
			<option value="0" <%if LinePC = "0" then response.write "selected"%> >No</option>
			</select>
		<input type="submit" name="setting" value="确定" />
	</form>
</div>
</br>
</center>
</body>
</html>
