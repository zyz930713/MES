<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#339966">
<center>
<form id="fmPack" method="post" action="PackingOperationSelect.asp" >
	<table width="700" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Select Packing Type ѡ���װ����</td>
      </tr>
	  <tr>
	  	<td><table border='0' cellpadding='0' cellspacing='5'><tr>
			<td ><input type="radio" name="rad_pack_type" value="NPI" <%if session("pack_type")="NPI" then response.Write("checked") end if%> >
&nbsp;NPI ��Ʒ </td>
	  	</tr>
		
			<tr>
			<td ><input type="radio" name="rad_pack_type" value="SCRAP" <%if session("pack_type")="SCRAP" then response.Write("checked") end if%> >
&nbsp;Scrap ����Ʒ </td>
			</tr>
            <tr>
			<td ><input type="radio" name="rad_pack_type" value="RWK" <%if session("pack_type")="SCRAP" then response.Write("checked") end if%> >
&nbsp;RWK Scrap ����Ʒ </td>
			</tr>
             <tr>
			<td ><input type="radio" name="rad_pack_type" value="RWKFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;RWK  Final Good ��Ʒ </td>
			</tr>
             <tr>
			<td ><input type="radio" name="rad_pack_type" value="TKBFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;TKB  Final Good ��Ʒ </td>
			</tr>
            <tr>
			<td ><input type="radio" name="rad_pack_type" value="TSDMarigoldFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;Marigold TSD Final Good ��Ʒ </td>
			</tr>
            <tr>
			<td ><input type="radio" name="rad_pack_type" value="TSDMarigoldTKB" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;Marigold TSD TKB ��Ʒ </td>
			</tr>
             <tr>
			<td ><input type="radio" name="rad_pack_type" value="TSDMarigoldRWK" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;Marigold TSD RWK ��Ʒ </td>
			</tr>
            <tr>
			<td ><input type="radio" name="rad_pack_type" value="TSDFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;TSD Final Good ��Ʒ </td>
			</tr>
            
             <tr>
			<td ><input type="radio" name="rad_pack_type" value="TSDTKBFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;TSD TKB Final Good ��Ʒ </td>
			</tr>
            <tr>
            	<td ><input type="radio" name="rad_pack_type" value="TSDRWKFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;TSD RWK Final Good ��Ʒ </td>
			</tr>	
            <tr><td ><input type="radio" name="rad_pack_type" value="SemiFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;�ֹ��� ��Ʒ </td>
			</tr>
            <tr><td ><input type="radio" name="rad_pack_type" value="HVFG" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;�Զ��� ��Ʒ </td>
			</tr>
            <tr><td ><input type="radio" name="rad_pack_type" value="Exception" <%if session("pack_type")="FG" then response.Write("checked") end if%> >
&nbsp;�����װ ��Ʒ </td>
			</tr>
			</table></td>
	  </tr>
	</table>
<br />
<input name="Next" type="button" value="Next ��һ��" onClick="fmPack.submit();" />
&nbsp;
<input name="Close" type="button" id="Close" onClick="window.close()" value="Close �ر�"> 
		  
</form>	
<table ><tr><td><a href="code_defect_code.asp">��ά�벹ȱ�ݴ���</a></td></tr></table>
</center>

</body>
</html>
