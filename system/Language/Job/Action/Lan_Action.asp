<script language="javascript">
var strReportTitle="Action Tracking|�������" 
var arrReportTitle=strReportTitle.split("|")
var strReportName="Report Name|��������" 
var arrReportName=strReportName.split("|")
var strReportPeriod="Close Time|����ʱ��" 
var arrReportPeriod=strReportPeriod.split("|")
var strFrom="From|��" 
var arrFrom=strFrom.split("|")
var strTo="To|��" 
var arrTo=strTo.split("|")
var strFactory="Factory|��" 
var arrFactory=strFactory.split("|")
var strType="Type|�û�" 
var arrType=strType.split("|")
var strThisStation="Station (Numerator)|վ�����ӣ�" 
var arrThisStation=strThisStation.split("|")
var strThisAction="Action (Numerator)|���裨���ӣ�" 
var arrThisAction=strThisAction.split("|")
var strReferStation="Station (Denominator)|վ����ĸ��" 
var arrReferStation=strReferStation.split("|")
var strReferAction="Action (Denominator)|���裨��ĸ��" 
var arrReferAction=strReferAction.split("|")
var strFamily="Family|����" 
var arrFamily=strFamily.split("|")
var strMailRecievers="Mail Recievers|�ռ���" 
var arrMailRecievers=strMailRecievers.split("|")
var strAvailableRecievers="Available Recievers|��ѡ�ռ���" 
var arrAvailableRecievers=strAvailableRecievers.split("|")
var strSelectedRecievers="Selected Recievers|��ѡ�ռ���" 
var arrSelectedRecievers=strSelectedRecievers.split("|")
var strGenerate="Generate|����" 
var arrGenerate=strGenerate.split("|")
var strReset="Reset|����" 
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