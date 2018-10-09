<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Rework Job Number|维修工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchReworkJobType="Rework Job Type|维修类型" 
var arrSearchReworkJobType=strSearchReworkJobType.split("|")
var strSearchStartTime="Rework Start Time|维修开始时间" 
var arrSearchStartTime=strSearchStartTime.split("|")
var strSearchEndTime="Rework End Time|维修结束时间" 
var arrSearchEndTime=strSearchEndTime.split("|")

var strSearchOperatorCode="Operator Code|操作工工号" 
var arrSearchOperatorCode=strSearchOperatorCode.split("|")

var strBrowse="Browse Rework Job Records (Default in past 7 days)|浏览维修工单记录（默认7天以内）" 
var arrBrowse=strBrowse.split("|")

var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")

var strJobNumber="Rework Job Number|维修工单号" 
var arrJobNumber=strJobNumber.split("|")

var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")

var strJobType="Rework Job Type|维修类型" 
var arrJobType=strJobType.split("|")

var strReworkQty="Rework Quantity|维修数量" 
var arrReworkQty=strReworkQty.split("|")

var strGoodQty="Good Quantity|合格数量" 
var arrGoodQtys=strGoodQty.split("|")

var strRejectQty="Reject Quantity|报废数量" 
var arrRejectQty=strRejectQty.split("|")

var strStartTime="Start Time|开始时间" 
var arrStartTime=strStartTime.split("|")

var strEndTime="End Time|结束时间" 
var arrEndTime=strEndTime.split("|")

var strOperatorCode="Operator Code|操作工工号" 
var arrOperatorCode=strOperatorCode.split("|")

var strYield="Yield|良率" 
var arrYield=strYield.split("|")

var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")

var strbtnUpdate="Update|更新" 
var arrbtnUpdate=strbtnUpdate.split("|")

var strUpdateTitle="Update Job Quantity|更新维修工单数量" 
var arrUpdateTitle=strUpdateTitle.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchReworkType.innerText=arrSearchReworkJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchStartTime.innerText=arrSearchStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchEndTime.innerText=arrSearchEndTime[<%=session("language")%>]}catch(e){}
try{inner_SearchOperatorCode.innerText=arrSearchOperatorCode[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Part_Number.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}

try{inner_JobType.innerText=arrJobType[<%=session("language")%>]}catch(e){}

try{inner_ReworkQty.innerText=arrReworkQty[<%=session("language")%>]}catch(e){}
try{inner_GoodQty.innerText=arrGoodQtys[<%=session("language")%>]}catch(e){}
try{inner_RejectQty.innerText=arrRejectQty[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_EndTime.innerText=arrEndTime[<%=session("language")%>]}catch(e){}
try{inner_OperatorCode.innerText=arrOperatorCode[<%=session("language")%>]}catch(e){}
try{inner_Yield.innerText=arrYield[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{document.all.btnUpdate.value=arrbtnUpdate[<%=session("language")%>]}catch(e){}
try{inner_UpdateTitle.innerText=arrUpdateTitle[<%=session("language")%>]}catch(e){}
}
</script>