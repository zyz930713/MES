<script language="javascript">
var strActionInfo="Action Info|������Ϣ" 
var arrActionInfo=strActionInfo.split("|")
var strDefectcodeInfo="Defectcode Info|ȱ�ݴ�����Ϣ" 
var arrDefectcodeInfo=strDefectcodeInfo.split("|")
var strUpdate="Update|��  ��"
var arrUpdate=strUpdate.split("|")
var strReset="Reset|��  ��"
var arrReset=strReset.split("|")
function language()
{
try{inner_ActionInfo.innerText=arrActionInfo[<%=session("language")%>]}catch(e){}
try{inner_DefectcodeInfo.innerText=arrDefectcodeInfo[<%=session("language")%>]}catch(e){}
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>