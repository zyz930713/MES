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
		
		 if (!document.form1.PART_NUMBER.value)
   		 {
    		alert("������12NC�ţ�");
			document.form1.PART_NUMBER.focus();
			return false;	
		
		}	
		
			
			
			if(document.form1.PART_NUMBER.value=="430307703161" || document.form1.PART_NUMBER.value=="430307703261"   || document.form1.PART_NUMBER.value=="430307705051"    )   //Hektor outer
			
			
			{
				document.form1.action="Outer_Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
			
				else if(document.form1.PART_NUMBER.value=="430307703151" || document.form1.PART_NUMBER.value=="430307703251" || document.form1.PART_NUMBER.value=="430307705041"  ) // Hektor inner
			
			
			{
				document.form1.action="inner_Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
				else if(document.form1.PART_NUMBER.value=="430307701182" || document.form1.PART_NUMBER.value=="430307703741" || document.form1.PART_NUMBER.value=="430307700394"  || document.form1.PART_NUMBER.value=="430307700393" || document.form1.PART_NUMBER.value=="430307704321") // HV
			
			
			{
				document.form1.action="HV_Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
				else if(document.form1.PART_NUMBER.value=="430307704701"       || document.form1.PART_NUMBER.value=="430307702931"      ) // HV
			
			
			{
				document.form1.action="FRANKLIN_Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
				else if(document.form1.PART_NUMBER.value=="430307703141" || document.form1.PART_NUMBER.value=="430307703241" ||  document.form1.PART_NUMBER.value=="430307705031")// Hektor POT ASSY 
			
			
			{
				document.form1.action="ALL_Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
			
			
			
				else if(document.form1.PART_NUMBER.value=="430307612751" || document.form1.PART_NUMBER.value=="430307615241" || document.form1.PART_NUMBER.value=="430307704331" || document.form1.PART_NUMBER.value=="430307700454" || document.form1.PART_NUMBER.value=="430307702721" || document.form1.PART_NUMBER.value=="430307703841" || document.form1.PART_NUMBER.value=="430307616111"   || document.form1.PART_NUMBER.value=="430307704401"   )
			
			
			{
				document.form1.action="Plate_Print2.asp";
			    document.form1.submit();
				
				
			}
			
			else if(document.form1.PART_NUMBER.value=="430307702921" ||  document.form1.PART_NUMBER.value=="430307705741"  )  //WINSLET POT ASSY
			
			
			{
				document.form1.action="Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
			
			else if(document.form1.PART_NUMBER.value=="430307704631" || document.form1.PART_NUMBER.value=="430307704841" || document.form1.PART_NUMBER.value=="430307705601" )  //Marigold pot assy
			
			
			{
				document.form1.action="Pot_Print2.asp";
			    document.form1.submit();
				
				
			}
			
			else
			
			{
				
				alert ("12NCû������")
				}
	}
	 
</script>
</head>
<body onLoad="form1.PACKED_USER.focus();" bgcolor="#339966">

<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">( Pot Assy Print)</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<Tr>
		<td height="30" colspan="2"><font color="#FF0000" size="20"><%=request.QueryString("word")%></font></td>
	</Tr>
	<%end if%>
    <tr>
      <td width="200" >Operator Code ����</td>
      <td width="430" ><input name="PACKED_USER" type="text" id="PACKED_USER" value=<%=PACKED_USER%> ></td>
    </tr>
    <tr>
      <td >Part Number �Ϻ�</td>
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