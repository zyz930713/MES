<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchCreateTime="Create Time|���ʱ��" 
var arrSearchCreateTime=strSearchCreateTime.split("|")
var strSearchJobCloseTime="Close Time|�ر�ʱ��" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strSearchPlaner="Planer|�ƻ���" 
var arrSearchPlaner=strSearchPlaner.split("|")
var strSearchProgress="Progress|����" 
var arrSearchProgress=strSearchProgress.split("|")
var strSearchProgressTypeSelect="-- Select Type --|-- ѡ������ --" 
var arrSearchProgressTypeSelect=strSearchProgressTypeSelect.split("|")
var strSearchProgressFinished="Finished|������" 
var arrSearchProgressFinished=strSearchProgressFinished.split("|")
var strSearchProgressProgressing="Progressing|������" 
var arrSearchProgressProgressing=strSearchProgressProgressing.split("|")

var strBrowse="Browse Master Job Records|�����������¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strProgress="Progress|������" 
var arrProgress=strProgress.split("|")
var strStock="Stock|���" 
var arrStock=strStock.split("|")
var strPlaner="Planer|�ƻ���" 
var arrPlaner=strPlaner.split("|")
var strStartQuantity="Start Quantity|�ƻ�����" 
var arrStartQuantity=strStartQuantity.split("|")
var strAssemblyGood="Assembly Good|���ߺϸ�Ʒ" 
var arrAssemblyGood=strAssemblyGood.split("|")
var strAssemblyYield="Assembly Yield|��������" 
var arrAssemblyYield=strAssemblyYield.split("|")
var strLineLost="Line Lost|������ʧ" 
var arrLineLost=strLineLost.split("|")
var strFinalGood="Final Good|���պϸ�Ʒ" 
var arrFinalGood=strFinalGood.split("|")
var strConfirmGood="Confirmed Good|ʵ�պϸ�Ʒ" 
var arrConfirmGood=strConfirmGood.split("|")
var strFinalYield="Final Yield|��������" 
var arrFinalYield=strFinalYield.split("|")
var strFinalScrap="Final Scrap|���ձ���" 
var arrFinalScrap=strFinalScrap.split("|")
var strFinalRemain="Remain|ʣ��" 
var arrFinalRemain=strFinalRemain.split("|")
var strInputTime="Input Time|Ͷ��ʱ��" 
var arrInputTime=strInputTime.split("|")
var strLastTime="Last Time|���ʱ��" 
var arrLastTime=strLastTime.split("|")
var strCycleTime="Cycle Time|��תʱ��" 
var arrCycleTime=strCycleTime.split("|")
var arrRecords=["No Records","û�м�¼"]
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchCreateTime.innerText=arrSearchCreateTime[<%=session("language")%>]}catch(e){}
try{inner_SearchJobCloseTime.innerText=arrSearchJobCloseTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchPlaner.innerText=arrSearchPlaner[<%=session("language")%>]}catch(e){}
try{inner_SearchProgress.innerText=arrSearchProgress[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[0].text=arrSearchProgressTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[1].text=arrSearchProgressFinished[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[2].text=arrSearchProgressProgressing[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Progress.innerText=arrProgress[<%=session("language")%>]}catch(e){}
try{inner_Stock.innerText=arrStock[<%=session("language")%>]}catch(e){}
try{inner_Planer.innerText=arrPlaner[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_AssemblyGood.innerText=arrAssemblyGood[<%=session("language")%>]}catch(e){}
try{inner_AssemblyYield.innerText=arrAssemblyYield[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_FinalGood.innerText=arrFinalGood[<%=session("language")%>]}catch(e){}
try{inner_ConfirmGood.innerText=arrConfirmGood[<%=session("language")%>]}catch(e){}
try{inner_FinalYield.innerText=arrFinalYield[<%=session("language")%>]}catch(e){}
try{inner_FinalScrap.innerText=arrFinalScrap[<%=session("language")%>]}catch(e){}
try{inner_FinalRemain.innerText=arrFinalRemain[<%=session("language")%>]}catch(e){}
try{inner_InputTime.innerText=arrInputTime[<%=session("language")%>]}catch(e){}
try{inner_LastTime.innerText=arrLastTime[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>