<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
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
<form action="/Admin/DefectCode/AddDefectCode1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td colspan="2" class="t-c-greenCopy">Add a New Defect Code </td>
</tr>
<tr>
  <td>Defect Code  <span class="red">*</span> </td>
  <td class="red">
      <div align="left">
        <input name="defectcode" type="text" id="defectcode">
      </div></td>
    </tr>
  <tr>
    <td>Defect Name  <span class="red">*</span> </td>
    <td><input name="defectname" type="text" id="defectname" size="50"></td>
  </tr>
  <tr>
    <td>Chinese Name  <span class="red">*</span></td>
    <td><input name="chinesename" type="text" id="chinesename" size="50"></td>
  </tr>
  <tr>
    <td>Belonged Factory <span class="red">*</span></td>
    <td><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
		<%FactoryRight ""%>
        <%= getFactory("OPTION","",factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20">Transaction Type</td>
    <td height="20"><select name="transaction" id="transaction">
        <option>Select Type</option>
        <option value="0">None</option>
        <option value="1">Rework</option>
        <option value="2">Scrap</option>
		<option value="3">Readust</option>
      </select>
    </td>
  </tr>
  <tr>
    <td>Applied Materials <span class="red">*</span> </td>
    <td><div align="center">
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
              <select name="fromitem1" size="10" multiple id="fromitem1">
			  <%FactoryRight "M."%>
				<%=getMaterial(true,"OPTION",factorywhereoutside,"","")%>
				</select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem1,document.form1.toitem1)"></div></td>
            <td rowspan="7"><div align="center">
              <select name="toitem1" size="10" multiple id="toitem1">
              </select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="8" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem1,document.form1.fromitem1)"></div></td>
            <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem1,document.form1.toitem1)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem1,document.form1.fromitem1)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem1)"> </div></td>
          </tr>
        </table>
    </div></td>
  </tr>
  <tr>
    <td>Belonged Part</td>
    <td><div align="left">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow"><div align="center">Available Parts </div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center">Selected Parts </div></td>
            <td><div align="center">&nbsp;</div></td>
          </tr>
          <tr>
            <td rowspan="7"><div align="center">
                <select name="fromitem2" size="10" multiple id="fromitem2">
                  <%FactoryRight "P."%>
                  <%=getPart(true,"OPTION","",factorywhereoutside," order by P.PART_NUMBER","","",rcount,"","","","","")%>
                </select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem2,document.form1.toitem2)"></div></td>
            <td rowspan="7"><div align="center">
                <select name="toitem2" size="10" multiple id="toitem2">
                                                </select>
            </div></td>
            <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem2)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem2,document.form1.fromitem2)"></div></td>
            <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem2)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem2,document.form1.toitem2)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem2)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem2,document.form1.fromitem2)"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem2)"> </div></td>
          </tr>
        </table>
    </div></td>
  </tr>
  <tr>
    <td>Belonged Station <span class="red">*</span> </td>
    <td>
      <div align="left">
        <select name="station" id="station">
		<option value="">-- Select --</option>
		<%FactoryRight "S."%>
		<%=getStation(true,"OPTION","",factorywhereoutside," order by STATION_NAME","","")%>
        </select>
      </div></td>
    </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
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