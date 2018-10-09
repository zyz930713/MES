<script language="javascript">
var strBrowse="Browse a Job|浏览工单" 
var arrBrowse=strBrowse.split("|")
var strSummary="Summary Info|摘要" 
var arrSummary=strSummary.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strDefectCodeInfo="Defect Code Info|缺陷代码信息" 
var arrDefectCodeInfo=strDefectCodeInfo.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strDefectCodeName="Defect Code Name|缺陷代码名称" 
var arrDefectCodeName=strDefectCodeName.split("|")
var strDefectCodeQuantity="Defect Code Quantity|缺陷代码数量" 
var arrDefectCodeQuantity=strDefectCodeQuantity.split("|")
var strClose="Close|关  闭"
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