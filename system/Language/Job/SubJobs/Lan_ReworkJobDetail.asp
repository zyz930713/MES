<script language="javascript">
var strBrowse="Browse a Job|�������" 
var arrBrowse=strBrowse.split("|")
var strSummary="Summary Info|ժҪ" 
var arrSummary=strSummary.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strDefectCodeInfo="Defect Code Info|ȱ�ݴ�����Ϣ" 
var arrDefectCodeInfo=strDefectCodeInfo.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strDefectCodeName="Defect Code Name|ȱ�ݴ�������" 
var arrDefectCodeName=strDefectCodeName.split("|")
var strDefectCodeQuantity="Defect Code Quantity|ȱ�ݴ�������" 
var arrDefectCodeQuantity=strDefectCodeQuantity.split("|")
var strClose="Close|��  ��"
var arrClose=strClose.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_Summary.innerText=arrSummary[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_DefectCodeInfo.innerText=arrDefectCodeInfo[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_DefectCodeName.innerText=arrDefectCodeName[<%=session("language")%>]}catch(e){}
try{inner_DefectCodeQuantity.innerText=arrDefectCodeQuantity[<%=session("language")%>]}catch(e){}
try{document.all.Close.value=arrClose[<%=session("language")%>]}catch(e){}
}
</script>