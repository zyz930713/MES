<script language="javascript">
var strSearch="Search Job|��ѯ����"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLineName="LineName|�߱�" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|����" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchOptionFactory="Factory|����" 
var arrSearchOptionFactory=strSearchOptionFactory.split("|")
var strSearchJobStartTime="Job Create Time|������ʼʱ��" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="to|��" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Rework Job list (Default in past 7 days)|���Rework�����б�Ĭ��7�����ڣ�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|�Ƴ�" 
var arrPartType=strPartType.split("|")
var strLine="Line|�߱�" 
var arrLine=strLine.split("|")
var strCreateTime="Create Time|����ʱ��" 
var arrCreateTime=strCreateTime.split("|")
var strQuantity="Quanity|����" 
var arrQuantity=strQuantity.split("|")
var strPrintTimes="Print Times|��ӡ����" 
var arrPrintTimes=strPrintTimes.split("|")
var strRecords="No Records|û�м�¼" 
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