<%
dbServer="localhost"
dbPort="3306"
dbDriver="{MySQL ODBC 5.2 Unicode Driver}" 'change this depending on your driver and mysql version
dbDatabase=""
dbUser=""
dbPassword=""

'DBName= "Server="&dbServer &"; Port="&dbPort &"; Driver="&dbDriver &"; Database="&dbDatabase &"; User="&dbUser &"; Pwd="&dbPassword &"; OPTION=3;"
DBName="SERVER="&dbServer &"; PORT="&dbPort &"; DRIVER="&dbDriver &"; DATABASE="&dbDatabase &"; UID="&dbUser &"; PWD="&dbPassword &"; OPTION=3;"
'response.write DBName

%>