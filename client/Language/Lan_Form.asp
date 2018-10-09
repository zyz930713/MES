<script language="javascript">
var strTitle="Create New NCMR Ticket|新建异常处理单"
var arrTitle=strTitle.split("|")
var strFactory ="Factory|专业厂" 
var arrFactory=strFactory.split("|")
var strCreator="User Code|申请人工号" 
var arrCreator=strCreator.split("|")
var strLine="Family|家族" 
var arrLine=strLine.split("|")
var strKeyWord="Key Word|关键字" 
var arrKeyWord=strKeyWord.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strModel="Model|型号" 
var arrModel=strModel.split("|")
var strStation="Station|工站" 
var arrStation=strStation.split("|")
var strProlem="Problem Description|问题描述" 
var arrProlem=strProlem.split("|")
var strUpload="Upload File|上传文件" 
var arrUpload=strUpload.split("|")
var strActionPerson="NCMR Owner|NCMR 负责人" 
var arrActionPerson=strActionPerson.split("|")
var strBackUp="Backup|备选" 
var arrBackUp=strBackUp.split("|")
var strDueDate="Happen Date|发生日期" 
var arrDueDate=strDueDate.split("|")
var strDotice="Notice Persons|通知相关人" 
var arrDotice=strDotice.split("|")
var strSimilar ="Similar Recodes|相似问题记录" 
var arrSimilar=strSimilar.split("|")
var strGreen="Green:Finished;|绿色：完成；" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|红色：进行中；" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|蓝色：未进行；" 
var arrBlue=strBlue.split("|")
var strNum="NO.|序列" 
var arrNum=strNum.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strNumber1 ="NCMR NO.|异常单编号" 
var arrNumber1=strNumber1.split("|")
var strKeyWord1="Key Word|关键字" 
var arrKeyWord1=strKeyWord1.split("|")
var strJobNumber1="Job Number|工单号" 
var arrJobNumber1=strJobNumber1.split("|")
var strModel1="Model|型号" 
var arrModel1=strModel1.split("|")
var strCreateTime1="Create Time|申请时间" 
var arrCreateTime1=strCreateTime1.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")

function language()
{
try{inner_Title.innerText=arrTitle[<%=session("language")%>]}catch(e){} 
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){} 
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_KeyWord.innerText=arrKeyWord[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_JobNumber1.innerText=arrJobNumber1[<%=session("language")%>]}catch(e){}
try{inner_Model.innerText=arrModel[<%=session("language")%>]}catch(e){}
try{inner_Model1.innerText=arrModel1[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Prolem.innerText=arrProlem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_BackUp.innerText=arrBackUp[<%=session("language")%>]}catch(e){} 
try{inner_ActionPerson.innerText=arrActionPerson[<%=session("language")%>]}catch(e){}
try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_Dotice.innerText=arrDotice[<%=session("language")%>]}catch(e){}
try{inner_Similar.innerText=arrSimilar[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Number1.innerText=arrNumber1[<%=session("language")%>]}catch(e){}
try{inner_KeyWord1.innerText=arrKeyWord1[<%=session("language")%>]}catch(e){}
try{inner_CreateTime1.innerText=arrCreateTime1[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}

}
</script>