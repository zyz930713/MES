<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/CheckABC/ABCCheck.asp" -->
<!--#include virtual="/WOCF/PVS_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
SolID=trim(request("SolID"))
paramSF=trim(request("paramSF"))
LLA=trim(request("LLA"))
ULA=trim(request("ULA"))
LLB=trim(request("LLB"))
ULB=trim(request("ULB"))
Product_Name=trim(request("Product_Name"))
Action=trim(request("Action"))


if Action="Del" then



conn.Execute("Delete from TempCheckSol  where SolID='"&SolID&"'")
word="Successfully Del CheckABC Config."
action="history.back()"


else
conn.Execute("update TempCheckSol  set paramSF='"&paramSF&"',LLA='"&LLA&"',ULA='"&ULA&"',LLB='"&LLB&"',ULB='"&ULB&"',Product_Name='"&Product_Name&"' where SolID='"&SolID&"'")
word="Successfully edit CheckABC Config."
action="location.href='"&beforepath&"'"

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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->