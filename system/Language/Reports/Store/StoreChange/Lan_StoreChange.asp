<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchChangeTime="Change Time|���ʱ��" 
var arrSearchChangeTime=strSearchChangeTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strSearchChangeCode="Change Code|�����" 
var arrSearchChangeCode=strSearchChangeCode.split("|")
var strSearchChangeType="Change Type|�������" 
var arrSearchChangeType=strSearchChangeType.split("|")
var strSearchChangeTypeSelect="-- Select Type --|-- ѡ������ --" 
var arrSearchChangeTypeSelect=strSearchChangeTypeSelect.split("|")
var strSearchChangeEdit="Edit|�޸�" 
var arrSearchChangeEdit=strSearchChangeEdit.split("|")
var strSearchChangeDelete="Delete|ɾ��" 
var arrSearchChangeDelete=strSearchChangeDelete.split("|")

var strBrowse="Browse Store Records|�������¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strChangeType="Change Type|�������" 
var arrChangeType=strChangeType.split("|")
var strChangeCode="Change Code|�����" 
var arrChangeCode=strChangeCode.split("|")
var strChangeReason="Change Reason|�������" 
var arrChangeReason=strChangeReason.split("|")
var strChangeTime="Change Time|���ʱ��" 
var arrChangeTime=strChangeTime.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strOldStoreCode="Old Code|�ɵ������" 
var arrOldStoreCode=strOldStoreCode.split("|")
var strOldInputQuantity="Old Input Quantity|�ɵ���������" 
var arrOldInputQuantity=strOldInputQuantity.split("|")
var strOldStoreQuantity="Old Store Quantity|�ɵ��������" 
var arrOldStoreQuantity=strOldStoreQuantity.split("|")
var strNewStoreQuantity="New Store Quantity|�µ��������" 
var arrNewStoreQuantity=strNewStoreQuantity.split("|")
var strOldStoreTime="Old Store Time|�ɵ����ʱ��" 
var arrOldStoreTime=strOldStoreTime.split("|")
var strOldStoreType="Old Store Type|�ɵ��������" 
var arrOldStoreType=strOldStoreType.split("|")
var strOldNote="Old Note|ע��" 
var arrOldNote=strOldNote.split("|")
var strSubJobs="SubJobs|ϸ�ֹ�����" 
var arrSubJobs=strSubJobs.split("|")
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
try{inner_OldStoreCode.innerText=arrOldStoreCode[<%=session("language")%>]}catch(e){}
try{inner_OldInputQuantity.innerText=arrOldInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_OldStoreQuantity.innerText=arrOldStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_NewStoreQuantity.innerText=arrNewStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_OldStoreTime.innerText=arrOldStoreTime[<%=session("language")%>]}catch(e){}
try{inner_OldStoreType.innerText=arrOldStoreType[<%=session("language")%>]}catch(e){}
try{inner_OldNote.innerText=arrOldNote[<%=session("language")%>]}catch(e){}
try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
}
</script>