<script language="javascript">
var strDefectCodeDistributionReport="Defect Code Distribution|ȱ�ݴ���ֲ�" 
var arrDefectCodeDistributionReport=strDefectCodeDistributionReport.split("|")
var strDefectCode="Defect Code|ȱ�ݴ���" 
var arrDefectCode=strDefectCode.split("|")
var strDate="Data|����"
var arrDate=strDate.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
function language()
{
try{inner_DefectCodeDistributionReport.innerText=arrDefectCodeDistributionReport[<%=session("language")%>]}catch(e){}
try{inner_DefectCode.innerText=arrDefectCode[<%=session("language")%>]}catch(e){}
try{inner_Date.innerText=arrDate[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
}
</script>
