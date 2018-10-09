<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetERPAccount.asp" -->
<!--#include virtual="/Functions/GetERPReason.asp" -->

<%
	MaterialPartNumber=request("txtMaterialPartNumber")
	Opcode=request("txtOpcode")
	ScrapQty=request("txtScrapQty")
	LineName=request("txtLineName")
	Comments=request("txtComments")
	ERP_Account=request("ERP_Account")
	ERP_Reason=request("ERP_Reason")
	ERP_Refer=request("ERP_Refer")
	if(request.QueryString("Action")="1") then
		if (MaterialPartNumber<>"" and Opcode<>"")then
			Set TypeLib = CreateObject("Scriptlet.TypeLib")
 			Guid = TypeLib.Guid
    
			SQLStr="SELECT * FROM PIECE_PARTS_SCRAP WHERE 1=2"
			rs.open SQLStr,conn,1,3
			rs.addnew
			rs("MATERIAL_PART_NUMBER")=MaterialPartNumber
			rs("SCRAP_OP_CODE")=Opcode
			rs("QTY")=ScrapQty
			rs("LINE_NAME")=LineName
			rs("COMMENTS")=Comments
			rs("SCRAPPING_ACCOUNT")=ERP_Account
			rs("SCRAPPING_REASON")=ERP_Reason
			rs("SCRAPPING_REFERENCE")=ERP_Refer
			rs("SCRAP_DATETIME")=now()
			rs("SCRAP_TYPE")="0"
			rs("IS_SENT_APPROVE")="0"
			rs("GUID")=Guid
			rs.update
			response.write("<script>alert('报废成功!')</script>")
		end if 
	end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>

<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>

<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<%if store_finished=true then%>
<script language="javascript">
parent.document.form1.Next.disabled=true;
</script>
<%end if%>
<script language="javascript">
 	function SaveData()
	{
		if(document.getElementById("txtMaterialPartNumber").value=="")
		{
			window.alert("请输入料号！");
			return;
		}
		
		if(document.getElementById("txtOpcode").value=="")
		{
			window.alert("请输入工号！");
			return;
		}
		if(document.getElementById("txtScrapQty").value=="")
		{
			window.alert("请输入报废数量！");
			return;
		}
		
		if(isNaN(document.getElementById("txtScrapQty").value))
		{
			window.alert("报废数量应该为数字！");
			return;
		}
		
		if(document.getElementById("txtLineName").value=="")
		{
			window.alert("请输入线别！");
			return;
		}
		if(document.getElementById("ERP_Account").value=="")
		{
			window.alert("请选择报废帐号！");
			return;
		}
		if(document.getElementById("ERP_Reason").value=="")
		{
			window.alert("请选择报废原因！");
			return;
		}
		if(document.getElementById("ERP_Refer").value=="")
		{
			window.alert("请选择报废Reference！");
			return;
		}
		
		form1.action="ManualScrapPieceParts.asp?Action=1";
		form1.submit();
	}
	
	function SearchData()
	{
		window.open("ManualScrapPiecePartsSearch.asp");
	}
</script>
</head>

<body>
<form method="post" name="form1" target="_self" >
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
    <td height="20" colspan="8" class="t-b-midautumn">Piece Parts Manual Scrap</td>
  </tr>
  
 <tr>
    <td height="20">料号<span class="red">*</span></td>
    <td height="20" colspan="6">
	<input type="text" name="txtMaterialPartNumber" id="txtMaterialPartNumber">
	</td>
  </tr>
  
   <tr>
    <td height="20">工号<span class="red">*</span></td>
    <td height="20" colspan="6">
	<input type="text" name="txtOpcode" id="txtOpcode">
	</td>
  </tr>
   
   <tr>
    <td height="20">报废数量<span class="red">*</span></td>
    <td height="20" colspan="6">
	<input type="text" name="txtScrapQty" id="txtScrapQty">
	</td>
  </tr>
  
   <tr>
    <td height="20">线别<span class="red">*</span></td>
    <td height="20" colspan="6">
	<input type="text" name="txtLineName" id="txtLineName">
	</td>
  </tr>
   <tr>
    <td height="20">备注</td>
    <td height="20" colspan="6">
	<input type="text" name="txtComments" id="txtComments" >
	</td>
  </tr>
  <tr>
    <td height="20">报废帐号<span class="red">*</span></td>
    <td height="20" colspan="6"><label>
      <select name="ERP_Account" id="ERP_Account">
	  <option value="">--选择帐号--</option>
	  <%=getERPAccount()%>
      </select>
	  <br><span id="errorinsertERPaccount" class="red"></span>
    </label></td>
  </tr>
  <tr>
    <td height="20">报废原因<span class="red">*</span></td>
    <td height="20" colspan="6"><label>
      <select name="ERP_Reason" id="ERP_Reason">
	   <option value="">--选择理由--</option>
	  <%=getERPReason()%>
      </select>
	  <br><span id="errorinsertERPreason" class="red"></span>
    </label></td>
  </tr>
  <tr>
    <td height="20">报废Reference<span class="red">*</span></td>
    <td height="20" colspan="6"><input name="ERP_Refer" type="text" id="ERP_Refer" onFocus="focushandler(this)" onBlur="blurhandler(this)" size="100"><br><span id="errorinsertERPrefer" class="red"></span></td>
  </tr>
    <tr>
	 <td height="20" colspan="7">
	<input type="button" name="btnSave" id="btnSave" value="保存" onclick="SaveData()"  class="t-b-Yellow">&nbsp;
	<input type="button" name="btnSearch" id="btnSearch" value="报废查询" onclick="SearchData()"  class="t-b-Yellow">
	</td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->