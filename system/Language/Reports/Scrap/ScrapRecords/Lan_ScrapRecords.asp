<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchScrapTime="Scrap Time|报废时间" 
var arrSearchScrapTime=strSearchScrapTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchCode="Code|报废人" 
var arrSearchCode=strSearchCode.split("|")

var strBrowse="Browse Scrap Records|浏览报废记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strCheck="Check|复查" 
var arrCheck=strCheck.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strCode="Code|报废人" 
var arrCode=strCode.split("|")
var strInputQuantity="Input Quantity|生产数量" 
var arrInputQuantity=strInputQuantity.split("|")
var strScrapQuantity="Scrap Quantity|报废数量" 
var arrScrapQuantity=strScrapQuantity.split("|")
var strOntimeYield="OntimeYield|即时良率" 
var arrOntimeYield=strOntimeYield.split("|")
var strScrapTime="Scrap Time|报废时间" 
var arrScrapTime=strScrapTime.split("|")
var strScrapType="Scrap Type|报废类型" 
var arrScrapType=strScrapType.split("|")
var strReason="Reason|理由" 
var arrReason=strReason.split("|")
var strFormID="Form ID|表单编号" 
var arrFormID=strFormID.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")
var strNote="Note|注释" 
var arrNote=strNote.split("|")
var arrERPJobStatus=["ERP Job Status","ERP工单状态"]
var arrScrapAccount=["Scrap Account","报废账号"]
var arrScrapReason=["Scrap Reason","报废原因"]
var arrScrapReference=["Scrap Reference","报废参照"]
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchScrapTime.innerText=arrSearchScrapTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Check.innerText=arrCheck[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_ScrapQuantity.innerText=arrScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_OntimeYield.innerText=arrOntimeYield[<%=session("language")%>]}catch(e){}
try{inner_ScrapTime.innerText=arrScrapTime[<%=session("language")%>]}catch(e){}
try{inner_ScrapType.innerText=arrScrapType[<%=session("language")%>]}catch(e){}
try{inner_Reason.innerText=arrReason[<%=session("language")%>]}catch(e){}
try{inner_FormID.innerText=arrFormID[<%=session("language")%>]}catch(e){}
try{inner_NoRecords.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_Note.innerText=arrNote[<%=session("language")%>]}catch(e){}
try{td_ERPJobStatus.innerText=arrERPJobStatus[<%=session("language")%>]}catch(e){}
try{td_ScrapAccount.innerText=arrScrapAccount[<%=session("language")%>]}catch(e){}
try{td_ScrapReason.innerText=arrScrapReason[<%=session("language")%>]}catch(e){}
try{td_ScrapReference.innerText=arrScrapReference[<%=session("language")%>]}catch(e){}
}
</script>