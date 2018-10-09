<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Admin/DefectCode/DefectCode.asp"
defectid=request("defectid")
defectcode=request("defectcode")
chinesename=request("chinesename")
belongpart=request("belongpart")
factory=request("factory")
part=request("part")
station=request("station")
defectNID=request("defectNID")
where=""
if defectNID<>"" then
where=where&" and D.NID='"&defectNID&"'"
end if
if defectid<>"" then
where=where&" and lower(D.DEFECT_CODE) like '%"&lcase(defectid)&"%'"
end if
if defectcode<>"" then
where=where&" and lower(D.DEFECT_NAME) like '%"&lcase(defectcode)&"%'"
end if
if chinesename<>"" then
where=where&" and lower(D.DEFECT_CHINESE_NAME) like '%"&lcase(chinesename)&"%'"
end if
if belongpart<>"" then
where=where&" and lower(APPLIED_PART_ID) like '%"&lcase(belongpart)&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and D.FACTORY_ID is null"
else
where=where&" and D.FACTORY_ID='"&factory&"'"
end if
if part="" or part="all" then
where=where&""
elseif part="null" then
where=where&" and D.APPLIED_PART_ID is null"
else
where=where&" and D.APPLIED_PART_ID='"&part&"'"
end if
if station="" or station="all" then
where=where&""
elseif station="null" then
where=where&" and D.STATION_ID is null"
else
where=where&" and D.STATION_ID='"&station&"'"
end if
pagepara="&factory="&factory&"&part="&part&"&station="&station&"&defectid="&defectid&"&defectcode="&defectcode&"&chinesename="&chinesename
FactoryRight "D."
SQL="select 1,D.*,F.FACTORY_NAME from DEFECTCODE D left join FACTORY F on D.FACTORY_ID=F.NID where 1=1 "&where&factorywhereoutsideand&" order by D.DEFECT_NAME"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Admin/Defectcode/Lan_Defectcode.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Admin/DefectCode/DefectCode.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
  	<td>DefectCode ID</td>
	<td><input name="defectNID" type="text" id="defectNID" value="<%=defectNID%>"></td>
    <td><span id="inner_SearchDefectID"></span></td>
    <td><input name="defectid" type="text" id="defectid" value="<%=defectid%>"></td>
    <td height="20"><span id="inner_SearchDefectCodeName"></span></td>
    <td><input name="defectcode" type="text" id="defectcode" value="<%=defectcode%>"></td>
    
  </tr>
  <tr>
  	<td height="20"><span id="inner_SearchChineseName"></span></td>
    <td><input name="chinesename" type="text" id="chinesename" value="<%=chinesename%>"></td>
    <td><span id="inner_SearchBelongedPart"></span></td>
    <td><input name="belongpart" type="text" id="belongpart" value="<%=belongpart%>"></td>
    <td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/DefectCode/AddDefectCode.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_Add"></span></a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="12"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
   
    <%if admin=true then%>
<td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">NID</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_DefectCode"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_DefectName"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ChineseName"></span></div></td>
  <td class="t-t-Borrow"><div align="center">
    <div align="center">
      <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value+'&part=<%=part%>&station=<%=station%>&defectcode=<%=defectcode%>&chinesename=<%=chinesename%>'">
        <option value="">Factory</option>
        <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
        <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
      </select>
    </div>
  </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Transaction"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_AppliedMaterials"></div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="part" id="part" onChange="location.href='<%=pagename%>?factory=<%=factory%>&part='+this.options[this.selectedIndex].value+'&station=<%=station%>&defectcode=<%=defectcode%>&chinesename=<%=chinesename%>'">
      <option value="" selected>Belonged Part</option>
      <option value="all" <%if part="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if part="null" then%>selected<%end if%>>None</option>
      <%FactoryRight "P."%>
      <%=getPart(true,"OPTION",part,factorywhereoutside," order by P.PART_NUMBER","","",rcount,"","","",null,null)%>
    </select>
  </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">
    <select name="station" id="station" onChange="location.href='<%=pagename%>?factory=<%=factory%>&part=<%=part%>&station='+this.options[this.selectedIndex].value+'&defectcode=<%=defectcode%>&chinesename=<%=chinesename%>'">
      <option value="">Belonged Station</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
	  <%FactoryRight "S."%>
      <%=getStation(true,"OPTION",station,factorywhereoutside," order by S.STATION_NAME","","")%>
    </select>
  </div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td width="39" height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
    <%if admin=true then%>
	
<td width="38" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:window.open('CopyDefectCode.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
<img src="/Images/IconCopy.gif" alt="Click to copy"></span></div></td>
    <td width="38" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:window.open ('EditDefectCode.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
<img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <!--<td width="47" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Action? If deleted, new job can not apply it forever. ')){window.open('DeleteDefectCode.asp?id=<%'=rs("NID")%>&path=<%'=path%>&query=<%'=query%>&defectcodename=<%'=rs("DEFECT_NAME")%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>-->
	<%end if%>
	<td width="181"><div align="center"><%= rs("NID")%></div></td>
    <td width="75" height="20"><div align="center"><%= rs("DEFECT_CODE")%></div></td>
    <td width="171" height="20"><div align="center"><%= rs("DEFECT_NAME")%></div></td>
    <td width="181"><div align="center"><%= rs("DEFECT_CHINESE_NAME")%>&nbsp;</div></td>
    <td width="181"><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td width="181"><div align="center">
	<%select case rs("TRANSACTION_TYPE")
	case "0"
	transaction_type="None"
	case "1"
	transaction_type="Rework"
	case "2"
	transaction_type="Scrap"
	end select%>
	<%=transaction_type%></div></td>
    <td width="181"><div align="center"><%if rs("MATERIAL_ID")<>"" then%><%=  getMaterial(true,"TEXT",""," where M.NID in ('"&replace(rs("MATERIAL_ID"),",","','")&"')"," ; ")%><%end if%>&nbsp;</div></td>
    <td width="181"><div align="center"><%if rs("APPLIED_PART_ID")<>"" then%>
	<%=  getPart(true,"TEXT",""," where P.NID in ('"&replace(rs("APPLIED_PART_ID"),",","','")&"')","",""," ; ",rcount,"","","","","")%><%end if%>&nbsp;</div></td>
    <td width="181" height="20"><div align="center"><%if rs("STATION_ID")<>"" then%><%=  getStation(true,"TEXT",""," where S.NID in ('"&replace(rs("STATION_ID"),",","','")&"')","",""," ; ")%><%end if%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="12"><div align="center">No records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
