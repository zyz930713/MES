<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
	response.Expires=0
	response.CacheControl="no-cache"
	pagename="Station.asp"
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>
<%
SQL = "select st.nid,st.station_name,st.station_chinese_name, sc.section_name from station_new st,section sc "
SQL = SQL+"where st.section_id = sc.nid and st.is_delete = 0 and st.transaction_type=0 and sc.status = 1 "
SQL = SQL+"order by sc.nid,st.wip_sequency"
rs.open SQL,conn,1,3
section = ""
%>

<body bgcolor="#339966">
<form action="Station1.asp" method="post" name="form1">
<table width="898" border="0" align="center"><tr><td width="831">
<table  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >
  <tr>
    <td height="20" colspan="2" class="t-t-DarkBlue"><div align="center">Select Station ѡ����վ</div></td>
  </tr>	
  <tr><td >
<%
i=1
stationCount=1
station=""
while not rs.eof 
  station = rs("nid") & "$" & rs("station_name")&"("& rs("station_chinese_name")&")"
  if rs("section_name") <> section then 
  	section = rs("section_name")
	if i>1 then
  		response.Write("</td></tr></table><tr><td>")
	end if
	i=1
	response.Write("<table border='0' cellpadding='0' cellspacing='5'><tr><td >"& section &"</td></tr><tr>")
  end if  
  response.Write("<td><input type='radio' name='radStationId' value='"&station&"' ")
  if Request.Cookies("current_station_id") = rs("nid") then 
  	response.Write("checked")
  end if
  response.Write(" >"& rs("station_name")&"("& rs("station_chinese_name")&")</td>")
  if (i mod 2) = 0 then
  	response.Write("</tr><tr>")
  end if
i = i+1
stationCount = stationCount+1
rs.movenext
wend 
if i>1 then
  response.Write("</td></tr></table>")
end if
%>
  </td></tr>  
</table>
</td></tr><tr><td align="center">

<input name="Next" type="submit" value="Next ��һ��" onClick="return formCheck()" />
&nbsp;
<input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="Close �ر�">
 

</td></tr>
<tr>
  <td height="38">&nbsp;</td>
</tr>
<tr><td><table width="846" height="179" border="1" align="center" cellpadding="1" class="today">	
 <tr>
   <td height="20" ><div align="center">ģ��˵��</div></td>
   <td ><div align="center">Ӧ������</div></td>
 </tr>
 <tr>
      <td width="339" height="20" ><div align="left">&nbsp;&nbsp;��Ʒ׷�ݺͿ���</div></td>
      <td width="491" ><a href="../PTC/PTC.asp" target="_blank">Product Tracking &amp; Control (��Ʒ׷�ݺͿ���)</a></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;��ӡ��ǩ����OQC���          &nbsp;&nbsp;</div></td>
      <td height="20" ><a href="PrintLabel.asp" target="_blank">Print Label(��ӡ��ǩ)</a></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;OQC�Բ�Ʒ���г��</div></td>
      <td height="20" ><a href="../OQC/OQCinfo.asp" target="_blank">OQC ����</a></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;���Ʒ����ƷԤ���</div></td>
      <td height="20" ><a href="../Store/PickJob.asp" target="_blank" >HFL Store ���</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="../Store/Store1.asp" target="_blank">PreStore Ԥ���</a></td>
    </tr>
    <tr>
      <td align="center"><div align="left">&nbsp;&nbsp;���������̽��а󶨷�����</div></td>
      <td ><a href="JobTrayLink.asp" target="_blank">Job Tray Link �������̰�</a></td>
    </tr>
    <tr>
      <td align="left">&nbsp;&nbsp;��Ʒ��װģ��&nbsp;</td>
      <td align="left"><a href="/Packing/SelectPackType.asp" target="_blank">Packing modules ��װģ��</a>&nbsp;&nbsp;   <a href="/ISSUE_RECORD/ISSUE.asp" target="_blank"> Reg ISSUE �쳣�����Ǽ�</a>  </td>
    </tr>
	
   
    
    <tr>
      <td height="20" >&nbsp;&nbsp;�ֿ�ģ��</td>
      <td height="20" ><a href="/Warehouse/stationscm.asp" target="_blank">WareHouse modules �ֿ�ģ��</a></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;Pot&amp;Membraned��ǩ��ӡ</div></td>
      <td height="20" ><p><a href="/EMR_Print/Pot_Print.asp" target="_blank">POT Assy</a> &nbsp;&nbsp;<a href="/EMR_Print/Mem_print.asp" target="_blank">Membrane Assy </a>&nbsp;&nbsp;<a href="/EMR_Print/Membrance_Print.asp" target="_blank">Maple Membrane Assy </a> &nbsp;&nbsp;&nbsp;</p>
        <p><a href="/EMR_Print/Frame_Print.asp" target="_blank">Maple Frame Assy </a></p></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;�ع����գ�RWK)&amp;���ˣ�TKB��</div></td>
      <td height="20" ><a href="/Rework/ReworkReceive.asp" target="_blank">Rework Receive �ع�����</a></td>
    </tr>
    <tr>
      <td height="20" ><div align="left">&nbsp;&nbsp;Check ABC</div></td>
      <td height="20" ><a href="/Packing/KEB_TSD_Maple.asp" target="_blank">Maple</a>&nbsp;&nbsp;&nbsp;<a href="/Packing/KEB_TSD_Pansy.asp" target="_blank">Pansy</a></td>
    </tr>
	<!--
	<tr>
      <td height="20" ><div align="center">
		  <input type="button" class="t-b-midautumn" value="Print Scrap List ��ӡ�����嵥" onClick="javascript:window.open('/Scrap/PrePrintScrapList.asp')">&nbsp;
          <input type="button" class="t-b-midautumn" value="Print Pack List ��ӡ��װ�嵥" onClick="javascript:window.open('/Store/PrintPackList.asp')">&nbsp;          
        </div></td>
    </tr>
	-->
  </table></td></tr></table>
<input type="hidden" id="computername" name="computername">
</form>	
</body>
<script language="javascript">
function formCheck()
{
	var isSltStation = false;
	for(var i=1;i<<%=stationCount%>;i++){
		if(form1.radStationId[i-1].checked){
			isSltStation = true;
			break;
		}
	}
	if(!isSltStation){
		alert("Please select one station! \n��ѡ��һ������վ��");
	}
	return isSltStation;
}
</script>


<script>
<%if Request.Cookies("computer_name") = "" then %>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
<%else%>
	document.form1.computername.value="<%=Request.Cookies("computer_name")%>"; 
<%end if%>
</script>



</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->