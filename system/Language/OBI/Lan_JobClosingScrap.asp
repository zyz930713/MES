<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchJobCloseTime="OBI Request Time|提交时间" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Job Auto (Scrap)|浏览工单自动报废" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strOBIResubmit="Resubmitable|可提交" 
var arrOBIResubmit=strOBIResubmit.split("|")
var strSelect="Select|选择" 
var arrSelect=strSelect.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strStartQuantity="Scrap Quantity|报废数量" 
var arrStartQuantity=strStartQuantity.split("|")
var strFinalGood="Final Good|最终合格品" 
var arrFinalGood=strFinalGood.split("|")
var strFinalYield="Final Yield|最终良率" 
var arrFinalYield=strFinalYield.split("|")
var strFinalScrap="Final Scrap|最终报废" 
var arrFinalScrap=strFinalScrap.split("|")
var strFinalRemain="Remain|剩余" 
var arrFinalRemain=strFinalRemain.split("|")
var strClosingTime="OBI Request Time|提交时间" 
var arrClosingTime=strClosingTime.split("|")
var strScrapAccount="Scrap Account|错误信息" 
var arrScrapAccount=strScrapAccount.split("|")
var strScrapReason="Scrap Reason|错误信息" 
var arrScrapReason=strScrapReason.split("|")
var strErrorCode="Error Code|错误代码" 
var arrErrorCode=strErrorCode.split("|")
var strError="Error Description|错误信息" 
var arrError=strError.split("|")
var strCheckAll="Check All|全选" 
var arrCheckAll=strCheckAll.split("|")
var strUnCheckAll="Uncheck All|全不选" 
var arrUnCheckAll=strUnCheckAll.split("|")
var strReSubmit="Re-submit|重新提交" 
var arrReSubmit=strReSubmit.split("|")
var strReset="Reset|重置" 
var arrReset=strReset.split("|")

function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchJobCloseTime.innerText=arrSearchJobCloseTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Select.innerText=arrSelect[<%=session("language")%>]}catch(e){}
try{inner_OBIResubmit.innerText=arrOBIResubmit[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_FinalGood.innerText=arrFinalGood[<%=session("language")%>]}catch(e){}
try{inner_FinalYield.innerText=arrFinalYield[<%=session("language")%>]}catch(e){}
try{inner_FinalScrap.innerText=arrFinalScrap[<%=session("language")%>]}catch(e){}
try{inner_FinalRemain.innerText=arrFinalRemain[<%=session("language")%>]}catch(e){}
try{inner_ClosingTime.innerText=arrClosingTime[<%=session("language")%>]}catch(e){}
try{inner_ScrapAccount.innerText=arrScrapAccount[<%=session("language")%>]}catch(e){}
try{inner_ScrapReason.innerText=arrScrapReason[<%=session("language")%>]}catch(e){}
try{inner_ErrorCode.innerText=arrErrorCode[<%=session("language")%>]}catch(e){}
try{inner_Error.innerText=arrError[<%=session("language")%>]}catch(e){}
try{document.all.CheckAll.value=arrCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.UncheckAll.value=arrUnCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.Resubmit.value=arrReSubmit[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>