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

var strBrowse="Browse Scrap Records|������ϼ�¼" 
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
var strOldScrapCode="Old Code|�ɵı�����" 
var arrOldScrapCode=strOldScrapCode.split("|")
var strOldScrapQuantity="Old Scrap Quantity|�ɵı�������" 
var arrOldScrapQuantity=strOldScrapQuantity.split("|")
var strNewScrapQuantity="New Scrap Quantity|�µı�������" 
var arrNewScrapQuantity=strNewScrapQuantity.split("|")
var strOldScrapTime="Old Scrap Time|�ɵı���ʱ��" 
var arrOldScrapTime=strOldScrapTime.split("|")
var strOldScrapType="Old Scrap Type|�ɵı�������" 
var arrOldScrapType=strOldScrapType.split("|")
var strOldNote="Old Note|ע��" 
var arrOldNote=strOldNote.split("|")
var strFormID="Form ID|�����" 
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