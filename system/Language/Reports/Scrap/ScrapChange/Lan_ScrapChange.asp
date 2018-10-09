<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchChangeTime="Change Time|变更时间" 
var arrSearchChangeTime=strSearchChangeTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchChangeCode="Change Code|变更人" 
var arrSearchChangeCode=strSearchChangeCode.split("|")
var strSearchChangeType="Change Type|变更类型" 
var arrSearchChangeType=strSearchChangeType.split("|")
var strSearchChangeTypeSelect="-- Select Type --|-- 选择类型 --" 
var arrSearchChangeTypeSelect=strSearchChangeTypeSelect.split("|")
var strSearchChangeEdit="Edit|修改" 
var arrSearchChangeEdit=strSearchChangeEdit.split("|")
var strSearchChangeDelete="Delete|删除" 
var arrSearchChangeDelete=strSearchChangeDelete.split("|")

var strBrowse="Browse Scrap Records|浏览报废记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strChangeType="Change Type|变更类型" 
var arrChangeType=strChangeType.split("|")
var strChangeCode="Change Code|变更人" 
var arrChangeCode=strChangeCode.split("|")
var strChangeReason="Change Reason|变更理由" 
var arrChangeReason=strChangeReason.split("|")
var strChangeTime="Change Time|变更时间" 
var arrChangeTime=strChangeTime.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strOldScrapCode="Old Code|旧的报废人" 
var arrOldScrapCode=strOldScrapCode.split("|")
var strOldScrapQuantity="Old Scrap Quantity|旧的报废数量" 
var arrOldScrapQuantity=strOldScrapQuantity.split("|")
var strNewScrapQuantity="New Scrap Quantity|新的报废数量" 
var arrNewScrapQuantity=strNewScrapQuantity.split("|")
var strOldScrapTime="Old Scrap Time|旧的报废时间" 
var arrOldScrapTime=strOldScrapTime.split("|")
var strOldScrapType="Old Scrap Type|旧的报废类型" 
var arrOldScrapType=strOldScrapType.split("|")
var strOldNote="Old Note|注释" 
var arrOldNote=strOldNote.split("|")
var strFormID="Form ID|表单编号" 
var arrFormID=strFormID.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeTime.innerText=arrSearchChangeTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeCode.innerText=arrSearchChangeCode[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeType.innerText=arrSearchChangeType[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[0].text=arrSearchChangeTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[1].text=arrSearchChangeEdit[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[2].text=arrSearchChangeDelete[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_ChangeType.innerText=arrChangeType[<%=session("language")%>]}catch(e){}
try{inner_ChangeCode.innerText=arrChangeCode[<%=session("language")%>]}catch(e){}
try{inner_ChangeReason.innerText=arrChangeReason[<%=session("language")%>]}catch(e){}
try{inner_ChangeTime.innerText=arrChangeTime[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_OldScrapCode.innerText=arrOldScrapCode[<%=session("language")%>]}catch(e){}
try{inner_OldScrapQuantity.innerText=arrOldScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_NewScrapQuantity.innerText=arrNewScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_OldScrapTime.innerText=arrOldScrapTime[<%=session("language")%>]}catch(e){}
try{inner_OldScrapType.innerText=arrOldScrapType[<%=session("language")%>]}catch(e){}
try{inner_OldNote.innerText=arrOldNote[<%=session("language")%>]}catch(e){}
try{inner_FormID.innerText=arrFormID[<%=session("language")%>]}catch(e){}
}
</script>