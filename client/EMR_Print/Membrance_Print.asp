<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">

	function SaveData()
	{
		//document.form1.computername.value="KES-IT-05";
	 	if(document.form1.computername.value=="")
		{
			window.alert("Can not get computer name of this machine,Please contact with IT. \nû���ܻ�ȡ��������������ϵIT��");
			return;
		}
		  
		 if (!document.form1.PACKED_USER.value)
   		 {
    		alert("�����빤�ţ�");
			document.form1.PACKED_USER.focus();
			return false;	
		
		}		
		
		
		else{
			
			document.form1.action="Membrance2_Print.asp";
			document.form1.submit();
		}
	}
	 
</script>
</head>
<body onLoad="form1.PACKED_USER.focus();" bgcolor="#339966">

<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(Membrance Assy Print)</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<Tr>
		<td height="20" colspan="2"><%=request.QueryString("word")%></td>
	</Tr>
	<%end if%>
    <tr>
      <td >Operator Code ����</td>
      <td ><input name="PACKED_USER" type="text" id="PACKED_USER" value=<%=PACKED_USER%> ></td>
    </tr>
    <tr>
      <td >�Ϻ�</td>
      <td ><input name="PART_NUMBER" type="text" id="PART_NUMBER" size="50" value=<%=PART_NUMBER%> ></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    
	
	<tr><td colspan="2">&nbsp;</td></tr>
    <tr>
      <td colspan="2" align="center">
	
	  <input type="hidden" id="computername" name="computername">
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Next ��һ��" onClick="SaveData();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close �ر�" onClick="javascript:window.close();">
      &nbsp;  
          
	  </td>
    </tr>
   <tr align="center"><td colspan="2">&nbsp;</td></tr>
  </form>
</table>
</body>
<script>
<%if Request.Cookies("computer_name") = "" then %>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
<%else%>
	document.form1.computername.value="<%=Request.Cookies("computer_name")%>"; 
<%end if%>
</script>
</html>