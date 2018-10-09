<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from DEFECTCODE where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/DefectCode/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/DefectCode/CopyDefectCode1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Copy a New Defect Code</td>
</tr>
<tr>
  <td width="116" height="20">Defect Code  <span class="red">*</span> </td>
  <td width="478" height="20" class="red">
      <div align="left">
        <input name="defectcode" type="text" id="defectcode" value="<%=rs("DEFECT_CODE")%>" size="50">
      </div></td>
    </tr>
  <tr>
    <td height="20">Defect Name  <span class="red">*</span> </td>
    <td height="20"><input name="defectname" type="text" id="defectname" value="<%=rs("DEFECT_NAME")%>" size="50"></td>
  </tr>
  <tr>
    <td height="20">Chinese Name <span class="red">*</span></td>
    <td height="20"><input name="chinesename" type="text" id="chinesename" value="<%=rs("DEFECT_CHINESE_NAME")%>" size="50"></td>
  </tr>
  <tr>
    <td height="20">Belonged Factory <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20">Applied Materials <span class="red">*</span> </td>
    <td height="20"><div align="center">
        <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow">
              <div align="center">Available Materials </div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center">Selected Materials </div></td>
            <td><div align="center">&nbsp;</div></td>
          </tr>
          <tr>
            <td rowspan="7"><div align="center">
              <select name="fromitem" size="15" id="fromitem">
                <%
			  FactoryRight "M."
			  if not isnull(rs("MATERIAL_ID")) or rs("MATERIAL_ID")<>"" then
			  where=" where M.NID not in ('"&replace(rs("MATERIAL_ID"),",","','")&"')"
			  end if%>
                <%=getMaterial(true,"OPTION","",where&factorywhereoutsideand,"")%>
              </select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
            <td rowspan="7"><div align="center">
              <select name="toitem" size="15" id="toitem">
                <%if not isnull(rs("MATERIAL_ID")) or rs("MATERIAL_ID")<>"" then
			  where=" where M.NID in ('"&replace(rs("MATERIAL_ID"),",","','")&"')"
			  end if%>
                <%=getMaterial(true,"OPTION","",where,"")%>
              </select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem)"></div></td>
            <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
          </tr>
        </table>
    </div></td>
  </tr>
  <tr>
    <td height="20">Belonged Part</td>
    <td height="20"><select name="part" id="part">
        <option value="">-- Select --</option>
        <%FactoryRight "P."%>
        <%=getPart(true,"OPTION",rs("APPLIED_PART_ID"),factorywhereoutside," order by P.PART_NUMBER","","",rcount,"","",tabi,a_SUB_ID,a_SUB_Number)%>
    </select></td>
  </tr>
  <tr>
    <td height="20">Belonged Station <span class="red">*</span> </td>
    <td height="20"><div align="left">
        <select name="station" id="station">
          <option value="">-- Select --</option>
          <%=getStation(true,"OPTION",rs("STATION_ID"),"","","","")%>
        </select>
    </div></td>
  </tr>
  
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="materialscount" type="hidden" id="materialscount">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Save">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
&nbsp;
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->