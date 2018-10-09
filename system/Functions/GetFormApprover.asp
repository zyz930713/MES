<%
function getFormApprover(applicant,first_role,second_role,third_role,fourth_role)
	 
ouput=""

getFormApprover=output

end function

function getApprover(role,applicant,flag)
set rskq=server.CreateObject("adodb.recordset")

Select case role
case "AR00000001": //Direct  Manager
		//检查主管工号是否为空
	SQLkq="select reportto from personnel01 where code='"&applicant&"'"
	rskq.open SQLkq,connkq,1,3
	if isnull(rskq("reportto")) then
	response.Redirect("/infoerror.asp?errorcode=E0001")
	end if
	rskq.close
	//申请人主管是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb 		where code=(select reportto from personnel01 where code='"&applicant&"'))"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	approvorcode=rskq("code")
	approvor=rskq("name")
	approvorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select reportto from personnel01 	where code='"&applicant&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	approvorcode=rskq("code")
	approvor=rskq("name")
	approvorenglish=rskq("englishname")
	end if
	else 
	output="未指定"
	end if
	output=approvor
	output1=approvorcode
	output2=approvorenglish

	rskq.close
	
case "AR00000002": //IT Infrastructure Supervisor 
	code="1203"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close
	
case "AR00000003":  //Department Manager
	//检查部门经理工号是否为空
	SQLkq="select code from keskq30.dbo.line where name='"&session("cdepartment")&"'"
	rskq.open SQLkq,connkq,1,3
	if isnull(rskq("Code")) then
	response.Redirect("/infoerror.asp?errorcode=E0001")
	end if
	rskq.close
	//申请人部门经理是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code=(select code from keskq30.dbo.line where name='"&session("cdepartment")&"'))"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select code from keskq30.dbo.line where name='"&session("cdepartment")&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close
	
case "AR00000004": //VP  
	code="1001"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close
	
case "AR00000005": //Training Manager   
	code="1626"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close	
	
case "AR00000006": //Regional HR Director  
	code="1244"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close
	
case "AR00000007":  //IT Manager
	//检查部门经理工号是否为空
	SQLkq="select code from keskq30.dbo.line where name='电脑部'"
	rskq.open SQLkq,connkq,1,3
	if isnull(rskq("Code")) then
	response.Redirect("/infoerror.asp?errorcode=E0001")
	end if
	rskq.close
	//IT部门经理是否有代理人
	'SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code=(select code from keskq30.dbo.line where name='"&session("cdepartment")&"'))"
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code=(select code from keskq30.dbo.line where name='电脑部'))"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select code from keskq30.dbo.line where name='电脑部')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close


case "AR00000008": //HR lobby
	//check the location of this applicant
	SQLkq="select Department from personnel01 where code='"&applicant&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then 
		if instr(rskq("Department"),"KES2")<>0 or instr(rskq("Department"),"Sisonic")<>0 then
			code="6128"
		else
			code="6124"
		end if
	end if
	rskq.close
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close
	
case "AR00000011": //GM
	//check the location of this applicant
	SQLkq="select Department from personnel01 where code='"&applicant&"'"
	'Response.Write(SQLkq)
	'Response.End()
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then 
		if instr(rskq("Department"),"KES2")<>0 or instr(rskq("Department"),"Sisonic")<>0 then
			code="1543"
		else
			code="1120"
		end if
	end if
	rskq.close
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close
	
case "AR00000012": //QA Lab Manager   
	code="1503"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close

case "AR00000013": //Finance Director   
	code="1510"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close
	
case "AR00000014": //HR General Affairs Specialist    
	code="1613"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close
	
case "AR00000015": //IT Device Administrator    
	code="1239"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close
	
case "AR00000016": //IT Device Administrator    
	code="1239"
	//approvor是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code='"&code&"')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code='"&code&"'"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish
	rskq.close
	
case "AR00000017": //IT Device Administrator    
	code="9999"
	
	output="系统"
	output1="9999"
	output2="System"

	//rskq.close

case "AR00000019":  //HR Manager
	//检查部门经理工号是否为空
	SQLkq="select code from keskq30.dbo.line where name='人力资源部'"
	rskq.open SQLkq,connkq,1,3
	if isnull(rskq("Code")) then
	response.Redirect("/infoerror.asp?errorcode=E0001")
	end if
	rskq.close
	//申请人部门经理是否有代理人
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select agent from personnelweb where code=(select code from keskq30.dbo.line where name='人力资源部'))"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	//有代理人
	agent=true
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	else
	agent=false
	end if
	rskq.close
	//没有代理人，赋值
	SQLkq="select code,convert(nvarchar,name) as name,englishname from personnel01 where code=(select code from keskq30.dbo.line where name='人力资源部')"
	rskq.open SQLkq,connkq,1,3
	if not rskq.eof then
	if agent=false then
	//无代理人
	departmentapprovorcode=rskq("code")
	departmentapprovor=rskq("name")
	departmentapprovorenglish=rskq("englishname")
	end if
	end if
	output=departmentapprovor
	output1=departmentapprovorcode
	output2=departmentapprovorenglish

	rskq.close

case else:
	output=""
	output1=""
	output2="None"

End select

getApprover=""
if flag=1 then
	getApprover=output
elseif flag=0 then
	getApprover=output1
else
	getApprover=output2
end if

set rskq=nothing

end function

%>