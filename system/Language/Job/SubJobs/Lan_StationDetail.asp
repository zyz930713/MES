<script language="javascript">
var strActionInfo="Action Info|动作信息" 
var arrActionInfo=strActionInfo.split("|")
var strDefectcodeInfo="Defectcode Info|缺陷代码信息" 
var arrDefectcodeInfo=strDefectcodeInfo.split("|")
var strUpdate="Update|更  新"
var arrUpdate=strUpdate.split("|")
var strReset="Reset|重  设"
var arrReset=strReset.split("|")
function language()
{
try{inner_ActionInfo.innerText=arrActionInfo[<%=session("language")%>]}catch(e){}
try{inner_DefectcodeInfo.innerText=arrDefectcodeInfo[<%=session("language")%>]}catch(e){}
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>