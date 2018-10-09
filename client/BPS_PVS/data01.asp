<!--#include file="conn.asp"-->
<%
dim ChartData,i
ChartData = ""
i = 1
set Rs = Server.CreateObject("adodb.recordset")
Sql = "EXEC [dbo].[Get_Charts_Data]"
Rs.open sql,Conn,1,1
if not rs.eof then
	while not rs.eof
		ListType = rs("ListType")
		DisplayName = rs("DisplayName")
		PrFOR = rs("PrFOR")
		PrSum = rs("PrSum")
			if ListType = "null" then
				ChartData = ChartData & "<point name='' y='' />"
				i = i + 1
			else
				ChartData = ChartData & "<point name='"& DisplayName &"' y='"& Formatnumber(PrFOR,2,0) &"' />"
			end if
		rs.movenext
	wend
end if
rs.close
set rs = nothing
%>

<?xml version="1.0" encoding="UTF-8"?>
<anychart>
	<charts>
		<chart plot_type="CategorizedHorizontal">
			<data_plot_settings default_series_type="Bar" >
				<bar_series point_padding="0.8" group_padding="0.8">
					<label_settings enabled="true" >
						<font family="Arial" color="" />
					</label_settings>
					<tooltip_settings enabled="True">
					</tooltip_settings>
				</bar_series>
			</data_plot_settings>
			
			<chart_settings>
				<chart_background enabled="false"/>
				<title enabled="false">
					<text>Sales of ACME Corp.</text>
				</title>
				<axes>
					<x_axis>
						<title enabled="false">
							<text>Retail Channel</text>
						</title>
					</x_axis>

					<y_axis>
						<title>
							<text>FOR</text>
						</title>
					</y_axis>
				</axes>
			</chart_settings>

			<data>
				<series name="FOR" type="Bar">
					<%=ChartData%>
				</series>
			</data>
		
		</chart>
	</charts>
</anychart>