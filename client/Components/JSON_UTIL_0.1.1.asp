<%
Function QueryToJSON(dbc, sql)
        Dim rs, jsa
        'Set rs = dbc.Execute(sql)
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,dbc,1,3
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
		rs.close
        Set QueryToJSON = jsa
End Function
%>