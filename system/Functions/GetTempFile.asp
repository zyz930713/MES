<%
	function getTempFile(myFileSystem)  
	dim tempFile,dotPos 
	tempFile=myFileSystem.getTempName 
	dotPos=instr(1,tempFile,".") 
	getTempFile=mid(tempFile,1,dotPos)&"xls" 
	end function
%>