<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<link href="/CSS/GeneralKES2.css" rel="stylesheet" type="text/css">
<html>
<head>
<meta http-equiv="refresh" content="5" charset=gb2312"/>

<title><%=application("SystemName")%></title>

</style>
</head>


<body bgcolor="#339966" >
<%


	 strWhere="where STATUS<>'Complete'"
	SQL="select * from Packing_Plan "&strWhere&" order by DELIVERY_TIME"
	rs.open sql,conn,1,3
	

%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >
	<tr>
        <td colspan="6" align="center" class="t-t-DarkBlue">Packing Plan ��װ�ƻ�</td>
    </tr>
	<tr align="center">
		<td >�ƻ�����</td>
		<td >�ͺ�</td>
		<td >����</td>
		<td >�ͻ�����</td>
		<td >��ע</td>
		<td >����ʱ��</td>
	</tr>
<%	

i=1
while not rs.eof
	result="Pass"
	'trColor="#00CC00"
	trColor="#FF0000"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rs("STATUS")="Wait" then
		
		
			'trColor="#FFFF00"
		else
			trColor="#FF0000"
  end if
	
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=rs("PLAN_DATE")%></td>
		<td><%=rs("PART_NUMBER")%></td>
		<td><%=rs("QUANTITY")%></td>
		<td><%=rs("CUSTOMER_NAME")%></td>
		<td><%=rs("REMARK")%>&nbsp;</td>
		<td><%=rs("DELIVERY_TIME")%>&nbsp;</td>
	</tr>	
<%
i=i+1
rs.movenext		
wend

%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->