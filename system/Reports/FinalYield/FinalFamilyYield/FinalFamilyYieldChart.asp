<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!-- #include virtual="/FXChart/Include/ChartFX.ASP.Core.inc" -->
<!-- #include virtual="/FXChart/Include/ChartFX.ASP.Borders.inc" -->
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/getFinalYieldChart.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOLE_Open.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<%
If Request("seriesgroup") = "" Then
seriesgroup = "OVERALL"
Else
seriesgroup = Request("seriesgroup")
End If
If Request("factory") = "" Then
factory = "FA00000002"
Else
factory = Request("factory")
End If
If Request("fromww") = "" Then
fromww = 1
Else
fromww = cint(Request("fromww"))
End If
If Request("toww") = "" Then
toww = cint(DatePart("ww", date()))
Else
toww = cint(Request("toww"))
End If
If Request("yearnumber") = "" Then
yearnumber = year(date())
Else
yearnumber = cint(Request("yearnumber"))
End If
If Request.Form("imagewidth") = "" Then
imagewidth = 640
Else
imagewidth = cint(Request.Form("imagewidth"))
End If
If Request.Form("imageheight") = "" Then
imageheight = 360
Else
imageheight = cint(Request.Form("imageheight"))
End If
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL = "Select FINAL_YIELD,FAMILY_TARGET_YIELD,INPUT_QUANTITY,OUTPUT_QUANTITY,CHART_WEEK,CHART_YEAR,CHART_MONTH from FAMILYYIELD_WEEK where FACTORY_ID='" & factory & "' and CHART_YEAR="&yearnumber&" and CHART_WEEK is not null and FAMILY_NAME='"&seriesgroup&"' and CHART_WEEK>=" & fromww & " and CHART_WEEK<=" & toww & " order by CHART_WEEK"
getFinalYieldChart seriesgroup,SQL,imagewidth,imageheight,filePath,rnd_key&".Jpeg"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Final Family Yield Chart</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="javascript">
function formcheck()
{
	with(document.form1)
	{
		if (fromww.selectedIndex==0||toww.selectedIndex==0)
		{
			alert("Week period must be set!");
			return false;
		}
		else
		{
			if (fromww.selectedIndex>=toww.selectedIndex)
			{
			alert("Start week is large than end week!");
			return false;
			}
		}
	}
}
</script>
</head>

<body>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-c-greenCopy">Final Family Yield Chart </td>
  </tr>
  <form id="form1" name="form1" method="post" action="" onSubmit="return formcheck()">
  <tr>
    <td>Family
      <select name="seriesgroup" id="seriesgroup">
	  <option value="">--Select--</option>
	  <option value="OVERALL" <%if seriesgroup="OVERALL" then%>selected<%end if%>>Overall</option>
	  <%=getSeriesGroup("CHARTOPTION",seriesgroup,"","","")%>
      </select>
      Week from
      <select name="fromww" id="fromww">
        <option value="">--Select--</option>
        <%for i=1 to 52%>
        <option value="<%=i%>" <%if fromww=i then%>selected<%end if%>><%=i%></option>
        <%next%>
      </select>
to Week
<select name="toww" id="toww">
  <option value="">--Select--</option>
  <%for i=1 to 52%>
  <option value="<%=i%>" <%if toww=i then%>selected<%end if%>><%=i%></option>
  <%next%>
</select>
in 
<input name="yearnumber" type="text" id="yearnumber" value="<%=yearnumber%>" size="4" />
year Image Size
<select name="imagewidth" id="imagewidth">
  <option value="800" <%if imagewidth=800 then%>selected<%end if%>>800</option>
  <option value="640" <%if imagewidth=640 then%>selected<%end if%>>640</option>
  <option value="480" <%if imagewidth=480 then%>selected<%end if%>>480</option>
</select>
X
<select name="imageheight" id="imageheight">
  <option value="600" <%if imageheight=600 then%>selected<%end if%>>600</option>
  <option value="480" <%if imageheight=480 then%>selected<%end if%>>480</option>
  <option value="360" <%if imageheight=360 then%>selected<%end if%>>360</option>
</select>
<input name="factory" type="hidden" id="factory" value="<%=factory%>" />
  <input name="Regenerate" type="button" id="Regenerate" value="Regenerate" onClick="javascript:document.form1.action='/Reports/FinalYield/FinalFamilyYield/FinalFamilyYieldChart.asp';document.form1.target='_self';document.form1.submit();"/>
  <input name="exportfamily" type="radio" value="0" checked="checked">
Current family
<input name="exportfamily" type="radio" value="1">
All families <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:if (document.form1.exportfamily[0].checked||document.form1.exportfamily[1].checked){document.form1.action='/Reports/FinalYield/FinalFamilyYield/FinalFamilyYieldChart_Export.asp';document.form1.target='_blank';document.form1.submit();}"><img src="/Images/EXCEL.gif" width="30" height="30" align="absmiddle"></span></td>
  </tr>
  </form>
  <tr>
    <td height="20" class="t-c-greenCopy">Chart Image </td>
  </tr>
  <tr>
    <td><div align="center"><img src="/Reports/Excel/<%=rnd_key&".Jpeg"%>" /></div></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->