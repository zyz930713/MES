<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchJobCloseTime="OBI Request Time|�ύʱ��" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Job Auto (Scrap)|��������Զ�����" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strOBIResubmit="Resubmitable|���ύ" 
var arrOBIResubmit=strOBIResubmit.split("|")
var strSelect="Select|ѡ��" 
var arrSelect=strSelect.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strStartQuantity="Scrap Quantity|��������" 
var arrStartQuantity=strStartQuantity.split("|")
var strFinalGood="Final Good|���պϸ�Ʒ" 
var arrFinalGood=strFinalGood.split("|")
var strFinalYield="Final Yield|��������" 
var arrFinalYield=strFinalYield.split("|")
var strFinalScrap="Final Scrap|���ձ���" 
var arrFinalScrap=strFinalScrap.split("|")
var strFinalRemain="Remain|ʣ��" 
var arrFinalRemain=strFinalRemain.split("|")
var strClosingTime="OBI Request Time|�ύʱ��" 
var arrClosingTime=strClosingTime.split("|")
var strScrapAccount="Scrap Account|������Ϣ" 
var arrScrapAccount=strScrapAccount.split("|")
var strScrapReason="Scrap Reason|������Ϣ" 
var arrScrapReason=strScrapReason.split("|")
var strErrorCode="Error Code|�������" 
var arrErrorCode=strErrorCode.split("|")
var strError="Error Description|������Ϣ" 
var arrError=strError.split("|")
var strCheckAll="Check All|ȫѡ" 
var arrCheckAll=strCheckAll.split("|")
var strUnCheckAll="Uncheck All|ȫ��ѡ" 
var arrUnCheckAll=strUnCheckAll.split("|")
var strReSubmit="Re-submit|�����ύ" 
var arrReSubmit=strReSubmit.split("|")
var strReset="Reset|����" 
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