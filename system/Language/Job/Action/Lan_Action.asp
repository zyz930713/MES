<script language="javascript">
var strReportTitle="Action Tracking|步骤跟踪" 
var arrReportTitle=strReportTitle.split("|")
var strReportName="Report Name|报告名称" 
var arrReportName=strReportName.split("|")
var strReportPeriod="Close Time|结束时间" 
var arrReportPeriod=strReportPeriod.split("|")
var strFrom="From|从" 
var arrFrom=strFrom.split("|")
var strTo="To|到" 
var arrTo=strTo.split("|")
var strFactory="Factory|厂" 
var arrFactory=strFactory.split("|")
var strType="Type|用户" 
var arrType=strType.split("|")
var strThisStation="Station (Numerator)|站（分子）" 
var arrThisStation=strThisStation.split("|")
var strThisAction="Action (Numerator)|步骤（分子）" 
var arrThisAction=strThisAction.split("|")
var strReferStation="Station (Denominator)|站（分母）" 
var arrReferStation=strReferStation.split("|")
var strReferAction="Action (Denominator)|步骤（分母）" 
var arrReferAction=strReferAction.split("|")
var strFamily="Family|家族" 
var arrFamily=strFamily.split("|")
var strMailRecievers="Mail Recievers|收件人" 
var arrMailRecievers=strMailRecievers.split("|")
var strAvailableRecievers="Available Recievers|可选收件人" 
var arrAvailableRecievers=strAvailableRecievers.split("|")
var strSelectedRecievers="Selected Recievers|已选收件人" 
var arrSelectedRecievers=strSelectedRecievers.split("|")
var strGenerate="Generate|产生" 
var arrGenerate=strGenerate.split("|")
var strReset="Reset|重置" 
var arrReset=strReset.split("|")
function language()
{
try{inner_ReportTitle.innerText=arrReportTitle[<%=session("language")%>]}catch(e){}
try{inner_ReportName.innerText=arrReportName[<%=session("language")%>]}catch(e){}
try{inner_ReportPeriod.innerText=arrReportPeriod[<%=session("language")%>]}catch(e){}
try{inner_From.innerText=arrFrom[<%=session("language")%>]}catch(e){}
try{inner_To.innerText=arrTo[<%=session("language")%>]}catch(e){}
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Type.innerText=arrType[<%=session("language")%>]}catch(e){}
try{inner_ThisStation.innerText=arrThisStation[<%=session("language")%>]}catch(e){}
try{inner_ThisAction.innerText=arrThisAction[<%=session("language")%>]}catch(e){}
try{inner_ReferStation.innerText=arrReferStation[<%=session("language")%>]}catch(e){}
try{inner_ReferAction.innerText=arrReferAction[<%=session("language")%>]}catch(e){}
try{inner_Family.innerText=arrFamily[<%=session("language")%>]}catch(e){}
try{inner_MailRecievers.innerText=arrMailRecievers[<%=session("language")%>]}catch(e){}
try{inner_AvailableRecievers.innerText=arrAvailableRecievers[<%=session("language")%>]}catch(e){}
try{inner_SelectedRecievers.innerText=arrSelectedRecievers[<%=session("language")%>]}catch(e){}
try{document.all.Generate.value=arrGenerate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>