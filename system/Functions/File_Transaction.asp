<%
function UploadFile(old_file_id,old_file_name,file_in,file_type,max_size,uploaded_file_name,uploaded_file_status,inventory_type)
	Set File = Upload.Files(file_in)
	If Not File Is Nothing Then
	file_name=File.FileName
	file_extension=File.Ext
	content_type=File.ContentType
	UploadFile=save_file(old_file_id,file,file_name,file_extension,content_type,uploaded_file_status,inventory_type)
	uploaded_file_name=file_name
	else
		if old_file_id<>"" then
		UploadFile=old_file_id
		uploaded_file_name=old_file_name
			if uploaded_file_status<>"" then
				set rsRS=server.CreateObject("adodb.recordset")
				SQLRS="update "&getInventoryDatabase(inventory_type)&" set STATUS="&uploaded_file_status&" where ID='"&old_file_id&"'"
				rsRS.open SQLRS,conn,1,3
				set rsRS=nothing
			end if
		else
		UploadFile=null
		uploaded_file_name=null
		end if
	end if
end function

function save_file(old_file_id,file,file_name,file_extension,content_type,uploaded_file_status,inventory_type)
	if old_file_id<>"" then
	SQLF="select * from "&getInventoryDatabase(inventory_type)&" where ID='"&old_file_id&"'"
	else
	SQLF="select * from "&getInventoryDatabase(inventory_type)&" where FILE_NAME='"&file_name&"' and CREATOR_CODE='"&session("code")&"'"
	end if
	rsF.open SQLF,conn,1,3
	if rsF.eof then
	rsF.addnew
	file_index=gen_key(20)&gen_key(20)
	NID="FI"&NID_SEQ("SYSTEM_FILE")
	rsF("ID")=NID
	rsF("FILE_INDEX")=file_index
	rsF("FILE_EXTENSION")=file_extension
	rsF("FILE_NAME")=file_name
	rsF("CONTENT_TYPE")=content_type
	if uploaded_file_status="0" then
	rsF("STATUS")=0
	else
	rsF("STATUS")=1
	end if
	rsF("CREATE_TIME")=now()
	rsF("CREATOR_CODE")=session("code")
	else
	NID=rsF("ID")
	file_index=rsF("FILE_INDEX")
	rsF("FILE_EXTENSION")=file_extension
	rsF("FILE_NAME")=file_name
	rsF("CONTENT_TYPE")=content_type
	rsF("UPDATE_TIME")=now()
	end if
	rsF.update
	rsF.close
	file.SaveAs application("LCD_Welcome")&"\"&file_index&file_extension '保存文件到服务器
end function
%>
