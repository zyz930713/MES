<script language="javascript">
var strDefectCodeInfo="Defect Code Info|ȱ�ݴ�����Ϣ" 
var arrDefectCodeInfo=strDefectCodeInfo.split("|")
var strNo="No|���" 
var arrNo=strNo.split("|")
var strJobNumber="Job Number|ϸ�ֹ�����" 
var arrJobNumber=strJobNumber.split("|")
var strStation="Station Name|վ��" 
var arrStation=strStation.split("|")
var strDefectName="Defect Code Name|ȱ�ݴ���" 
var arrDefectName=strDefectName.split("|")
var strDefectQuantity="Defect Code Quantity|ȱ������" 
var arrDefectQuantity=strDefectQuantity.split("|")
var strStepQuantity="Step Quantity|С������" 
var arrStepQuantity=strStepQuantity.split("|")
var strTotal="Total|�ϼ�"
var arrTotal=strTotal.split("|")
var strClose="Close|�ر�"
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