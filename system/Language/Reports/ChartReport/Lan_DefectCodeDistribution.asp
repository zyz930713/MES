<script language="javascript">
var strDefectCodeDistributionReport="Defect Code Distribution|缺陷代码分布" 
var arrDefectCodeDistributionReport=strDefectCodeDistributionReport.split("|")
var strDefectCode="Defect Code|缺陷代码" 
var arrDefectCode=strDefectCode.split("|")
var strDate="Data|日期"
var arrDate=strDate.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
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
