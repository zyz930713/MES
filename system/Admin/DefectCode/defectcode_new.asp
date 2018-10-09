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
pagename="/Admin/DefectCode/DefectCode_New.asp"
defectid=request("defectid")
defectcode=request("defectcode")
thisstatus=request.QueryString("status")
chinesename=request("chinesename")
belongpart=request("belongpart")
factory=request("factory")
part=request("part")
station=request("station")
where=""
if defectid<>"" then
where=where&" and lower(D.DEFECT_CODE) like '%"&lcase(defectid)&"%'"
end if
if defectcode<>"" then
where=where&" and lower(D.DEFECT_NAME) like '%"&lcase(defectcode)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and D.STATUS="&thisstatus
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
pagepara="&status="&thisstatus&"&factory="&factory&"&part="&part&"&station="&station&"&defectid="&defectid&"&defectcode="&defectcode&"&chinesename="&chinesename
FactoryRight "D."
SQL="select D.*,F.FACTORY_NAME from DEFECTCODE_New D left join FACTORY F on D.FACTORY_ID=F.NID where IS_DELETE=0 "&where&factorywhereoutsideand&" order by D.DEFECT_code"
rs.open SQL,conn,1,3
'response.Write(SQL)
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
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language();language_page()">
<form action="/Admin/DefectCode/DefectCode_New.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td><span id="inner_SearchDefectID"></span></td>
    <td><input name="defectid" type="text" id="defectid" value="<%=defectid%>"></td>
    <td height="20"><span id="inner_SearchDefectCodeName"></span></td>
    <td><input name="defectcode" type="text" id="defectcode" value="<%=defectcode%>"></td>
    <td height="20"><span id="inner_SearchChineseName"></span></td>
    <td><input name="chinesename" type="text" id="chinesename" value="<%=chinesename%>"></td>
   
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/DefectCode/AddDefectCode_New.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_Add"></span></a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <%if admin=true then%>
<td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  
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
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td  height="20" align="center"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>

    <td  height="20" align="center" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:window.open ('EditDefectCode_New.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
<img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
<td  height="20" align="center" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if(confirm('Are you sure to delete this Defect Code?\n您确定要删除此缺陷代码吗？')){window.open('DeleteDefectCodeNew.asp?id=<%=rs("NID")%>&actionname=<%=rs("DEFECT_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
    <td  height="20" align="center"><div align="center"><%= rs("DEFECT_CODE")%></div></td>
    <td  height="20" align="center"><div align="center"><%= rs("DEFECT_NAME")%></div></td>
    <td  align="center"><div align="center"><%= rs("DEFECT_CHINESE_NAME")%>&nbsp;</div></td>
    <td  align="center"><div align="center"><%= rs("FACTORY_NAME")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="13"><div align="center">No records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
