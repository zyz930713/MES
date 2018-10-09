<script language="javascript">
var strBrowse="Personal information|个人基本信息"
var arrBrowse=strBrowse.split("|")
var strCode="User Code|工号" 
var arrCode=strCode.split("|")
var strName="User Name|姓名" 
var arrName=strName.split("|")
var strChineseName="Chinese Name|中文姓名" 
var arrChineseName=strChineseName.split("|")
var strBackup="Backup Person|代理人" 
var arrBackup=strBackup.split("|")
var strLanguage="Language|语言" 
var arrLanguage=strLanguage.split("|")
var strEmail="Email|电子邮件" 
var arrEmail=strEmail.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strYieldAlertModel="Yield Alert By Model|按型号良率提醒" 
var arrYieldAlertModel=strYieldAlertModel.split("|")
var strYieldAlertLine="Yield Alert By Line|按线别良率提醒" 
var arrYieldAlertLine=strYieldAlertLine.split("|")
var strYieldAlertFamily="Yield Alert By Family|按家族良率提醒" 
var arrYieldAlertFamily=strYieldAlertFamily.split("|")
var strAlertNo="No|编号" 
var arrAlertNo=strAlertNo.split("|")
var strAlertModelPrefix="Model Prefix|型号前缀" 
var arrAlertModelPrefix=strAlertModelPrefix.split("|")
var strAlertLineName="Line|线别" 
var arrAlertLineName=strAlertLineName.split("|")
var strAlertFamilyName="Family|家族" 
var arrAlertFamilyName=strAlertFamilyName.split("|")
var strAlertFamilyNameOption="-- Family --|-- 家族 --" 
var arrAlertFamilyNameOption=strAlertFamilyNameOption.split("|")
var strAlertModelYield="Treshold|阀值" 
var arrAlertModelYield=strAlertModelYield.split("|")

var strRole="Role|角色" 
var arrRole=strRole.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_Logon.innerText=arrLogon[<%=session("language")%>]}catch(e){} 
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_Name.innerText=arrName[<%=session("language")%>]}catch(e){}
try{inner_ChineseName.innerText=arrChineseName[<%=session("language")%>]}catch(e){}
try{inner_Backup.innerText=arrBackup[<%=session("language")%>]}catch(e){}
try{inner_Language.innerText=arrLanguage[<%=session("language")%>]}catch(e){}
try{inner_Email.innerText=arrEmail[<%=session("language")%>]}catch(e){}
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_YieldAlertModel.innerText=arrYieldAlertModel[<%=session("language")%>]}catch(e){}
try{inner_YieldAlertLine.innerText=arrYieldAlertLine[<%=session("language")%>]}catch(e){}
try{inner_YieldAlertFamily.innerText=arrYieldAlertFamily[<%=session("language")%>]}catch(e){}
try{inner_AlertNo1.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo2.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo3.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo4.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo5.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo6.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo7.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo8.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertNo9.innerText=arrAlertNo[<%=session("language")%>]}catch(e){}
try{inner_AlertModelPrefix1.innerText=arrAlertModelPrefix[<%=session("language")%>]}catch(e){}
try{inner_AlertModelPrefix2.innerText=arrAlertModelPrefix[<%=session("language")%>]}catch(e){}
try{inner_AlertModelPrefix3.innerText=arrAlertModelPrefix[<%=session("language")%>]}catch(e){}
try{inner_AlertLineName4.innerText=arrAlertLineName[<%=session("language")%>]}catch(e){}
try{inner_AlertLineName5.innerText=arrAlertLineName[<%=session("language")%>]}catch(e){}
try{inner_AlertLineName6.innerText=arrAlertLineName[<%=session("language")%>]}catch(e){}
try{inner_AlertFamilyName7.innerText=arrAlertFamilyName[<%=session("language")%>]}catch(e){}
try{inner_AlertFamilyName8.innerText=arrAlertFamilyName[<%=session("language")%>]}catch(e){}
try{inner_AlertFamilyName9.innerText=arrAlertFamilyName[<%=session("language")%>]}catch(e){}
try{document.all.family1.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family2.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family3.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family4.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family5.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family6.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family7.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family8.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family9.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family10.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family11.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family12.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family13.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family14.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{document.all.family15.options[0].text=arrAlertFamilyNameOption[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield1.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield2.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield3.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield4.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield5.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield6.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield7.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield8.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_AlertModelYield9.innerText=arrAlertModelYield[<%=session("language")%>]}catch(e){}
try{inner_Role.innerText=arrRole[<%=session("language")%>]}catch(e){}
}
language()
</script>