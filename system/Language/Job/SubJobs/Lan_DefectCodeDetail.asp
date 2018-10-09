<script language="javascript">
var strDefectCodeInfo="Defect Code Info|缺陷代码信息" 
var arrDefectCodeInfo=strDefectCodeInfo.split("|")
var strNo="No|编号" 
var arrNo=strNo.split("|")
var strJobNumber="Job Number|细分工单号" 
var arrJobNumber=strJobNumber.split("|")
var strStation="Station Name|站名" 
var arrStation=strStation.split("|")
var strDefectName="Defect Code Name|缺陷代码" 
var arrDefectName=strDefectName.split("|")
var strDefectQuantity="Defect Code Quantity|缺陷数量" 
var arrDefectQuantity=strDefectQuantity.split("|")
var strStepQuantity="Step Quantity|小计数量" 
var arrStepQuantity=strStepQuantity.split("|")
var strTotal="Total|合计"
var arrTotal=strTotal.split("|")
var strClose="Close|关闭"
var arrClose=strClose.split("|")
function language()
{
try{inner_DefectCodeInfo.innerText=arrDefectCodeInfo[<%=session("language")%>]}catch(e){}
try{inner_No.innerText=arrNo[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_DefectName.innerText=arrDefectName[<%=session("language")%>]}catch(e){}
try{inner_DefectQuantity.innerText=arrDefectQuantity[<%=session("language")%>]}catch(e){}
try{inner_StepQuantity.innerText=arrStepQuantity[<%=session("language")%>]}catch(e){}
try{inner_Total.innerText=arrTotal[<%=session("language")%>]}catch(e){}
try{document.all.Close.value=arrClose[<%=session("language")%>]}catch(e){}
}
</script>