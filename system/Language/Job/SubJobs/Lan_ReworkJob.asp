<script language="javascript">
var strSearch="Search Job|查询工单"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLineName="LineName|线别" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|工厂" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchOptionFactory="Factory|工厂" 
var arrSearchOptionFactory=strSearchOptionFactory.split("|")
var strSearchJobStartTime="Job Create Time|工单开始时间" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="to|到" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Rework Job list (Default in past 7 days)|浏览Rework工单列表（默认7天以内）" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|制程" 
var arrPartType=strPartType.split("|")
var strLine="Line|线别" 
var arrLine=strLine.split("|")
var strCreateTime="Create Time|创建时间" 
var arrCreateTime=strCreateTime.split("|")
var strQuantity="Quanity|数量" 
var arrQuantity=strQuantity.split("|")
var strPrintTimes="Print Times|打印次数" 
var arrPrintTimes=strPrintTimes.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchOptionFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_CreateTime.innerText=arrCreateTime[<%=session("language")%>]}catch(e){}
try{inner_Quantity.innerText=arrQuantity[<%=session("language")%>]}catch(e){}
try{inner_PrintTimes.innerText=arrPrintTimes[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>